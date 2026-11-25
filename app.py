import os
import re
from tkinter import filedialog, messagebox
import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD

# Set up CustomTkinter appearance
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class ProduceProcessorApp(ctk.CTk, TkinterDnD.DnDWrapper):

  def __init__(self):
    super().__init__()

    # Initialize TkinterDnD
    self.TkdndVersion = TkinterDnD._require(self)

    # Window Configuration
    self.title("Produce Data Processor")
    self.geometry("760x1040")
    self.resizable(False, False)

    # Region configurations: Display Name -> Expected Region Token in Filename
    self.regions = [
        ("Alberta", "ALBERTA"),
        ("Manitoba", "MANITOBA"),
        ("L-Region", "L-REGION"),
        ("Saskatchewan", "SASKATCHEWAN"),
    ]

    # Storage for uploaded files
    # Each region maps to a dict: {"xlsx": path, "xls": path, "csv": path}
    self.region_files = {
        name: {"xlsx": None, "xls": None, "csv": None}
        for name, _ in self.regions
    }
    self.ad_final_file = None

    # UI References for dynamic updates
    self.region_ui_elements = {}
    self.ad_status_label = None

    self.create_widgets()

  def create_widgets(self):
    # Main Title Label
    title_label = ctk.CTkLabel(
        self,
        text="Produce Data Processor",
        font=ctk.CTkFont(size=24, weight="bold"),
    )
    title_label.pack(pady=(20, 10))

    # Scrollable frame for clean layout
    main_frame = ctk.CTkScrollableFrame(
        self, width=710, height=900, fg_color="transparent"
    )
    main_frame.pack(fill="both", expand=True, padx=20, pady=10)

    # Section Header: Regions
    region_header = ctk.CTkLabel(
        main_frame,
        text="Regional Produce Files (Max 3 files per zone: XLSX, XLS, CSV)",
        font=ctk.CTkFont(size=15, weight="bold"),
    )
    region_header.pack(anchor="w", pady=(0, 8))

    # 1. Four Regional Upload Zones
    for display_name, region_token in self.regions:
      self.create_region_zone(main_frame, display_name, region_token)

    # Separator
    sep = ctk.CTkFrame(main_frame, height=2, fg_color="#333333")
    sep.pack(fill="x", pady=18)

    # Section Header: Master AD File
    ad_header = ctk.CTkLabel(
        main_frame,
        text="Master AD Produce File (Required)",
        font=ctk.CTkFont(size=15, weight="bold"),
    )
    ad_header.pack(anchor="w", pady=(0, 8))

    # 2. Fifth Upload Zone (AD XX Produce Final Keyed & proofed.xlsx)
    self.create_ad_zone(main_frame)

    # 3. Start Processing Button
    self.process_btn = ctk.CTkButton(
        main_frame,
        text="Start Processing",
        font=ctk.CTkFont(size=15, weight="bold"),
        height=45,
        fg_color="#1f6aa5",
        hover_color="#144870",
        command=self.on_start_processing,
    )
    self.process_btn.pack(fill="x", pady=(25, 10))

  def create_region_zone(self, parent, display_name, region_token):
    frame = ctk.CTkFrame(
        parent,
        fg_color="#2b2b2b",
        corner_radius=8,
        height=145,
    )
    frame.pack(fill="x", pady=6)
    frame.pack_propagate(False)

    # Top Row: Region Title + Buttons
    top_row = ctk.CTkFrame(frame, fg_color="transparent")
    top_row.pack(fill="x", padx=15, pady=(10, 4))

    title_lbl = ctk.CTkLabel(
        top_row,
        text=display_name,
        font=ctk.CTkFont(size=15, weight="bold"),
    )
    title_lbl.pack(side="left")

    clear_btn = ctk.CTkButton(
        top_row,
        text="Clear",
        width=60,
        height=28,
        fg_color="#444444",
        hover_color="#555555",
        command=lambda d=display_name: self.clear_region(d),
    )
    clear_btn.pack(side="right", padx=(5, 0))

    browse_btn = ctk.CTkButton(
        top_row,
        text="Browse...",
        width=85,
        height=28,
        command=lambda d=display_name, r=region_token: self.browse_region_files(
            d, r
        ),
    )
    browse_btn.pack(side="right")

    # Bottom Row: Status indicators for the 3 slots
    slots_frame = ctk.CTkFrame(frame, fg_color="#222222", corner_radius=6)
    slots_frame.pack(fill="both", expand=True, padx=15, pady=(2, 10))

    xlsx_lbl = ctk.CTkLabel(
        slots_frame,
        text=f"1. WK##_PRODUCE_{region_token}.XLSX:  Missing",
        font=ctk.CTkFont(size=12),
        text_color="gray",
        anchor="w",
    )
    xlsx_lbl.pack(fill="x", padx=10, pady=(5, 1))

    xls_lbl = ctk.CTkLabel(
        slots_frame,
        text="2. Price Book .xls file:  Missing",
        font=ctk.CTkFont(size=12),
        text_color="gray",
        anchor="w",
    )
    xls_lbl.pack(fill="x", padx=10, pady=1)

    csv_lbl = ctk.CTkLabel(
        slots_frame,
        text="3. Produce .csv file:  Missing",
        font=ctk.CTkFont(size=12),
        text_color="gray",
        anchor="w",
    )
    csv_lbl.pack(fill="x", padx=10, pady=(1, 5))

    self.region_ui_elements[display_name] = {
        "xlsx": xlsx_lbl,
        "xls": xls_lbl,
        "csv": csv_lbl,
        "token": region_token,
    }

    # Register drag & drop across all child widgets in the zone
    drop_widgets = (
        frame,
        top_row,
        title_lbl,
        slots_frame,
        xlsx_lbl,
        xls_lbl,
        csv_lbl,
    )
    for w in drop_widgets:
      w.drop_target_register(DND_FILES)
      w.dnd_bind(
          "<<Drop>>",
          lambda e, d=display_name, r=region_token: self.handle_region_drop(
              e.data, d, r
          ),
      )

  def create_ad_zone(self, parent):
    frame = ctk.CTkFrame(
        parent,
        fg_color="#2b2b2b",
        corner_radius=8,
        height=80,
    )
    frame.pack(fill="x", pady=6)
    frame.pack_propagate(False)

    label = ctk.CTkLabel(
        frame,
        text="AD XX Produce Final Keyed & proofed.xlsx",
        font=ctk.CTkFont(size=14, weight="bold"),
    )
    label.pack(side="left", padx=15)

    self.ad_status_label = ctk.CTkLabel(
        frame,
        text="Drag & drop or browse file",
        font=ctk.CTkFont(size=12, slant="italic"),
        text_color="gray",
    )
    self.ad_status_label.pack(side="right", padx=15)

    browse_btn = ctk.CTkButton(
        frame,
        text="Browse...",
        width=90,
        height=32,
        command=self.browse_ad_file,
    )
    browse_btn.pack(side="right", padx=5)

    for w in (frame, label, self.ad_status_label):
      w.drop_target_register(DND_FILES)
      w.dnd_bind("<<Drop>>", lambda e: self.handle_ad_drop(e.data))

  # -------------------------------------------------------------------------
  # Validation & File Handling for Regions
  # -------------------------------------------------------------------------
  def process_region_files(self, file_paths, display_name, region_token):
    if len(file_paths) > 3:
      messagebox.showerror(
          "Too Many Files",
          f"You can upload a maximum of 3 files to '{display_name}'.",
      )
      return

    # Regex for WK##_PRODUCE_<REGION>.XLSX (case-insensitive, ## = any alphanumeric/text)
    expected_xlsx_pattern = re.compile(
        rf"^wk.+_produce_{re.escape(region_token)}\.xlsx$", re.IGNORECASE
    )

    for raw_path in file_paths:
      file_path = raw_path.strip('{}""\'')
      if not os.path.exists(file_path):
        messagebox.showerror("Error", f"File does not exist:\n{file_path}")
        continue

      file_name = os.path.basename(file_path)
      ext = os.path.splitext(file_name)[1].lower()

      if ext == ".xlsx":
        if not expected_xlsx_pattern.match(file_name):
          messagebox.showerror(
              "Invalid XLSX Filename",
              f"Invalid file for {display_name}:\n'{file_name}'\n\n"
              f"Expected format: WK##_PRODUCE_{region_token}.XLSX",
          )
          continue
        self.region_files[display_name]["xlsx"] = file_path

      elif ext == ".xls":
        self.region_files[display_name]["xls"] = file_path

      elif ext == ".csv":
        self.region_files[display_name]["csv"] = file_path

      else:
        messagebox.showerror(
            "Invalid File Type",
            f"Unsupported file '{file_name}'.\nOnly .xlsx, .xls, and .csv files"
            " are accepted.",
        )

    self.refresh_region_ui(display_name)

  def refresh_region_ui(self, display_name):
    files = self.region_files[display_name]
    ui = self.region_ui_elements[display_name]
    token = ui["token"]

    if files["xlsx"]:
      ui["xlsx"].configure(
          text=f"1. {os.path.basename(files['xlsx'])}", text_color="#2ecc71"
      )
    else:
      ui["xlsx"].configure(
          text=f"1. WK##_PRODUCE_{token}.XLSX:  Missing", text_color="gray"
      )

    if files["xls"]:
      ui["xls"].configure(
          text=f"2. Price Book .xls file:  {os.path.basename(files['xls'])}",
          text_color="#2ecc71",
      )
    else:
      ui["xls"].configure(
          text="2. Price Book .xls file:  Missing", text_color="gray"
      )

    if files["csv"]:
      ui["csv"].configure(
          text=f"3. Produce .csv file:  {os.path.basename(files['csv'])}",
          text_color="#2ecc71",
      )
    else:
      ui["csv"].configure(
          text="3. Produce .csv file:  Missing", text_color="gray"
      )

  def clear_region(self, display_name):
    self.region_files[display_name] = {"xlsx": None, "xls": None, "csv": None}
    self.refresh_region_ui(display_name)

  def browse_region_files(self, display_name, region_token):
    paths = filedialog.askopenfilenames(
        title=f"Select Files for {display_name}",
        filetypes=[
            ("Allowed Files", "*.xlsx *.xls *.csv"),
            ("All Files", "*.*"),
        ],
    )
    if paths:
      self.process_region_files(list(paths), display_name, region_token)

  def handle_region_drop(self, raw_data, display_name, region_token):
    files = self.tk.splitlist(raw_data)
    if files:
      self.process_region_files(list(files), display_name, region_token)

  # -------------------------------------------------------------------------
  # Validation & File Handling for AD Final File
  # -------------------------------------------------------------------------
  def process_ad_file(self, file_paths):
    if len(file_paths) > 1:
      messagebox.showerror(
          "Single File Only",
          "Please upload 1 and only 1 file for 'AD XX Produce Final Keyed &"
          " proofed.xlsx'.",
      )
      return

    file_path = file_paths[0].strip('{}""\'')
    if not os.path.exists(file_path):
      messagebox.showerror("Error", f"File does not exist:\n{file_path}")
      return

    file_name = os.path.basename(file_path)

    # Matches: AD ## - Produce Final Keyed&Proofed.xlsx (allows optional spaces around '-' and '&')
    ad_pattern = re.compile(
        r"^ad\s*.+\s*-\s*produce final keyed\s*&\s*proofed\.xlsx$",
        re.IGNORECASE,
    )

    if not ad_pattern.match(file_name):
      messagebox.showerror(
          "Invalid File",
          f"Invalid file:\n'{file_name}'\n\n"
          "Expected format: 'AD ## - Produce Final Keyed&Proofed.xlsx'",
      )
      return

    self.ad_final_file = file_path
    self.ad_status_label.configure(text=file_name, text_color="#2ecc71")

  def browse_ad_file(self):
    path = filedialog.askopenfilename(
        title="Select AD Produce Final Keyed & Proofed File",
        filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
    )
    if path:
      self.process_ad_file([path])

  def handle_ad_drop(self, raw_data):
    files = self.tk.splitlist(raw_data)
    if files:
      self.process_ad_file(list(files))

  # -------------------------------------------------------------------------
  # Start Processing Validation Check
  # -------------------------------------------------------------------------
  def on_start_processing(self):
    # 1. Check AD ## - Produce Final Keyed&Proofed file is uploaded
    if not self.ad_final_file:
      messagebox.showwarning(
          "Missing Required File",
          "Please upload the 'AD ## - Produce Final Keyed&Proofed.xlsx' file"
          " before starting.",
      )
      return

    # 2. Check that at least one of the 4 region zones has all 3 expected files
    completed_regions = []
    partial_regions = []

    for region_name, files in self.region_files.items():
      uploaded_count = sum(1 for v in files.values() if v is not None)
      if uploaded_count == 3:
        completed_regions.append(region_name)
      elif 0 < uploaded_count < 3:
        partial_regions.append(region_name)

    if not completed_regions:
      msg = (
          "At least one region zone must be completely populated with all 3"
          " expected files (WK##_PRODUCE_<REGION>.XLSX, Price Book .xls, and"
          " Produce .csv)."
      )
      if partial_regions:
        msg += f"\n\nIncomplete regions detected: {', '.join(partial_regions)}"
      messagebox.showwarning("Incomplete Regional Data", msg)
      return

    if partial_regions:
      messagebox.showwarning(
          "Incomplete Region Detected",
          "The following region(s) have missing files and cannot be processed"
          f" partially:\n{', '.join(partial_regions)}\n\nPlease complete or"
          " clear them before proceeding.",
      )
      return

    # Validation passed — ready for processing logic
    messagebox.showinfo(
        "Validation Passed",
        "Ready to process!\n\n"
        f"AD Master File: {os.path.basename(self.ad_final_file)}\n"
        f"Populated Regions: {', '.join(completed_regions)}",
    )


if __name__ == "__main__":
  app = ProduceProcessorApp()
  app.mainloop()