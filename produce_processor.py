import xlwings as xw
from xlwings.utils import rgb_to_int


class ProduceExcelProcessor:

  @staticmethod
  def _convert_columns_to_number(ws, columns, start_row=2):
    """Converts string-formatted numeric cells in the given columns into real numbers."""
    last_row = ws.used_range.last_cell.row
    if last_row < start_row:
      return

    for col in columns:
      col_range = ws.range(f"{col}{start_row}:{col}{last_row}")
      col_range.number_format = "General"

      values = col_range.value
      if not isinstance(values, list):
        values = [values]

      converted = []
      changed = False
      for val in values:
        if isinstance(val, str):
          cleaned = val.strip()
          if cleaned:
            try:
              num = float(cleaned)
              val = int(num) if num.is_integer() else num
              changed = True
            except ValueError:
              pass
        converted.append([val])

      if changed:
        col_range.value = converted

  @staticmethod
  def _clone_first_sheet_into(excel_app, target_wb, source_path, new_tab_name):
    """Clones the first worksheet of source_path into target_wb safely using a visible anchor."""
    # Delete any existing tab matching new_tab_name (case-insensitive)
    for sheet in list(target_wb.sheets):
      if new_tab_name.lower() in sheet.name.lower():
        sheet.delete()

    # Find the last VISIBLE sheet to use as a valid COM Copy anchor
    visible_sheets = [s for s in target_wb.sheets if s.visible]
    anchor_sheet = visible_sheets[-1] if visible_sheets else target_wb.sheets[0]
    anchor_sheet.visible = True

    before_count = len(target_wb.sheets)

    wb_src = excel_app.books.open(source_path)
    try:
      ws_src = wb_src.sheets[0]
      ws_src.api.Copy(After=anchor_sheet.api)

      # Verify a new sheet was actually added to target_wb
      if len(target_wb.sheets) > before_count:
        cloned_sheet = target_wb.sheets[anchor_sheet.index]
        cloned_sheet.name = new_tab_name
      else:
        # Fallback if Excel blocked cross-format (.xls -> .xlsx) sheet cloning
        new_sheet = target_wb.sheets.add(
            name=new_tab_name, after=anchor_sheet
        )
        src_last = ws_src.used_range.last_cell
        if (
            src_last.row > 1
            or src_last.column > 1
            or ws_src.range("A1").value is not None
        ):
          ws_src.range((1, 1), (src_last.row, src_last.column)).copy(
              destination=new_sheet.range("A1")
          )
          excel_app.api.CutCopyMode = False
    finally:
      wb_src.close()

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

    # Step 6b: Ensure columns E and H in "Current Week" are stored as numbers, not strings
    ProduceExcelProcessor._convert_columns_to_number(ws_current, ["E", "H"])

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

    # Step 11b: L-Region specific formula redirect on Cost & Sell
    if "l-region" in wb.name.lower():
      ProduceExcelProcessor.redirect_formula_update(wb)

    # Step 12: Touch-ups (reorder tabs and apply tab colors)
    ProduceExcelProcessor.touch_ups(wb)

  @staticmethod
  def update_sabw_price_book(excel_app, wb, xls_path):
    """Clones the SABW Price Book worksheet directly to preserve layout, widths, and barcodes."""
    ProduceExcelProcessor._clone_first_sheet_into(
        excel_app, wb, xls_path, "SABW Price Book"
    )

  @staticmethod
  def update_weekly_produce_ad(excel_app, wb, ad_xlsx_path):
    """Clones the Weekly Produce Ad worksheet directly to preserve spacing, formatting, and styles."""
    ProduceExcelProcessor._clone_first_sheet_into(
        excel_app, wb, ad_xlsx_path, "Weekly Produce Ad"
    )

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
  def redirect_formula_update(wb):
    """L-Region only: Updates J3 on 'Cost & Sell' and propagates down to the bottom."""
    try:
      ws_cost_sell = wb.sheets["Cost & Sell"]
    except Exception as e:
      raise ValueError(
          f"Could not find sheet 'Cost & Sell' in {wb.name}."
      ) from e

    ws_cost_sell.activate()
    ws_cost_sell.range("J3").formula = (
        '=IFERROR(_xlfn.LET(x,VLOOKUP(B3,\'SABW Price Book\'!E:Q,13,FALSE),'
        'IF(x=0,"",x)),"")'
    )

    last_row = ws_cost_sell.used_range.last_cell.row
    if last_row > 3:
      ws_cost_sell.range(f"J3:J{last_row}").api.FillDown()

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