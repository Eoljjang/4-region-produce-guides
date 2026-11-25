import xlwings as xw
from xlwings.utils import rgb_to_int


class ProduceExcelProcessor:

  @staticmethod
  def process_region_workbook(excel_app, wb, csv_path, xls_path, ad_xlsx_path):
    """Orchestrates all processing steps on WK##_PRODUCE_<REGION>.xlsx."""
    target_tabs = ["Current Week", "Previous Week", "Calculations"]

    # Step 2: Unhide required tabs
    for tab_name in target_tabs:
      try:
        sheet = wb.sheets[tab_name]
        sheet.visible = True
      except Exception as e:
        raise ValueError(
            f"Could not find or unhide sheet '{tab_name}' in {wb.name}."
        ) from e

    ws_current = wb.sheets["Current Week"]
    ws_prev = wb.sheets["Previous Week"]
    ws_calc = wb.sheets["Calculations"]

    # Step 3: Delete everything inside "Previous Week"
    ws_prev.clear()

    # Step 4: Copy all content & formatting from "Current Week" to "Previous Week", then clear "Current Week"
    last_cell = ws_current.used_range.last_cell
    if (
        last_cell.row > 1
        or last_cell.column > 1
        or ws_current.range("A1").value is not None
    ):
      source_range = ws_current.range((1, 1), (last_cell.row, last_cell.column))
      source_range.copy(destination=ws_prev.range("A1"))
      excel_app.api.CutCopyMode = False
      ws_current.clear()

    # Step 5: In "Previous Week", sort Column "O" from Smallest to Largest (vertically)
    prev_used = ws_prev.used_range
    if prev_used is not None and prev_used.rows.count > 1:
      prev_used.api.Sort(
          Key1=ws_prev.range("O1").api,
          Order1=1,  # 1 = xlAscending (Smallest to Largest)
          Header=1,  # 1 = xlYes (keeps Row 1 headers fixed at the top)
          Orientation=1,  # 1 = xlSortColumns (sorts rows top-to-bottom)
      )

    # Step 6: Open associated .csv file and paste its contents into "Current Week"
    wb_csv = excel_app.books.open(csv_path)
    try:
      ws_csv = wb_csv.sheets[0]
      csv_last = ws_csv.used_range.last_cell
      if (
          csv_last.row > 1
          or csv_last.column > 1
          or ws_csv.range("A1").value is not None
      ):
        csv_range = ws_csv.range((1, 1), (csv_last.row, csv_last.column))
        csv_range.copy(destination=ws_current.range("A1"))
        excel_app.api.CutCopyMode = False
    finally:
      wb_csv.close()

    # Step 7: Fix formulas in "Calculations" tab and propagate to the bottom
    ws_calc.range("A2").formula = "='Previous Week'!O2"
    ws_calc.range("E2").formula = (
        '=IF(ISBLANK(\'Previous Week\'!K2),"",\'Previous Week\'!K2)'
    )

    last_row = max(
        ws_calc.used_range.last_cell.row,
        ws_prev.used_range.last_cell.row,
    )

    if last_row > 2:
      ws_calc.range(f"A2:A{last_row}").api.FillDown()
      ws_calc.range(f"E2:E{last_row}").api.FillDown()

    # Step 8: Hide the tabs again
    for tab_name in target_tabs:
      wb.sheets[tab_name].visible = False

    # Step 9: Update SABW price book
    ProduceExcelProcessor.update_sabw_price_book(excel_app, wb, xls_path)

    # Step 10: Update Weekly Produce Ad
    ProduceExcelProcessor.update_weekly_produce_ad(excel_app, wb, ad_xlsx_path)

    # Step 11: Update date formula on Cost & Sell
    ProduceExcelProcessor.update_date(wb)

    # Step 12: Touch-ups (reorder tabs and apply tab colors)
    ProduceExcelProcessor.touch_ups(wb)

  @staticmethod
  def update_sabw_price_book(excel_app, wb, xls_path):
    """Clones the SABW Price Book worksheet directly to preserve layout, widths, and barcodes."""
    # 1. Delete the existing tab that CONTAINS "SABW Price Book" in the name
    for sheet in list(wb.sheets):
      if "sabw price book" in sheet.name.lower():
        sheet.delete()

    # 2. Open the uploaded "SABW Price Book.xls" file and clone the entire sheet into wb
    wb_xls = excel_app.books.open(xls_path)
    try:
      ws_xls = wb_xls.sheets[0]
      ws_xls.api.Copy(After=wb.sheets[-1].api)
    finally:
      wb_xls.close()

    # 3. Rename the newly cloned sheet to "SABW Price Book"
    cloned_sheet = wb.sheets[-1]
    cloned_sheet.name = "SABW Price Book"

  @staticmethod
  def update_weekly_produce_ad(excel_app, wb, ad_xlsx_path):
    """Clones the Weekly Produce Ad worksheet directly to preserve spacing, formatting, and styles."""
    # 1. Delete the existing tab named "Weekly Produce Ad" (or containing it, case-insensitive)
    for sheet in list(wb.sheets):
      if "weekly produce ad" in sheet.name.lower():
        sheet.delete()

    # 2. Open the uploaded "AD XX Produce Final Keyed & proofed.xlsx" file and clone its sheet into wb
    wb_ad = excel_app.books.open(ad_xlsx_path)
    try:
      ws_ad = wb_ad.sheets[0]
      ws_ad.api.Copy(After=wb.sheets[-1].api)
    finally:
      wb_ad.close()

    # 3. Rename the newly cloned sheet to "Weekly Produce Ad"
    cloned_sheet = wb.sheets[-1]
    cloned_sheet.name = "Weekly Produce Ad"

  @staticmethod
  def update_date(wb):
    """Sets Cell B1 on the 'Cost & Sell' tab to reference 'Weekly Produce Ad'!B4."""
    try:
      ws_cost_sell = wb.sheets["Cost & Sell"]
    except Exception as e:
      raise ValueError(
          f"Could not find sheet 'Cost & Sell' in {wb.name}."
      ) from e

    ws_cost_sell.activate()
    ws_cost_sell.range("B1").formula = "='Weekly Produce Ad'!B4"

  @staticmethod
  def touch_ups(wb):
    """Moves 'SABW Price Book' and 'Weekly Produce Ad' to the first two tab positions and colors them blue."""
    ws_sabw = wb.sheets["SABW Price Book"]
    ws_ad = wb.sheets["Weekly Produce Ad"]

    # Move SABW Price Book to position 1, and Weekly Produce Ad directly after it (position 2)
    ws_sabw.api.Move(Before=wb.sheets[0].api)
    ws_ad.api.Move(After=ws_sabw.api)

    # Set both worksheet tab colors to standard Excel blue using COM Tab.Color
    blue_color = rgb_to_int((0, 112, 192))
    ws_sabw.api.Tab.Color = blue_color
    ws_ad.api.Tab.Color = blue_color

    # Activate the first tab so the workbook opens cleanly on tab 1
    ws_sabw.activate()