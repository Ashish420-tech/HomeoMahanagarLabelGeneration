import sys
import os
import json
import logging
import pandas as pd
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QCompleter, QTableWidgetItem, QMessageBox
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase.pdfmetrics import stringWidth
import win32api
import win32print
from pathlib import Path
import platform
import wikipedia
import requests
from bs4 import BeautifulSoup

# ---- DPI awareness for Windows ----
if platform.system() == "Windows":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass

def get_system_scaling():
    try:
        app = QtWidgets.QApplication.instance()
        if app is None:
            app = QtWidgets.QApplication(sys.argv)
        screen = app.primaryScreen()
        if not screen:
            return 1.0
        screen_size = screen.size()
        base_width, base_height = 1920, 1080
        scale_w = screen_size.width() / base_width
        scale_h = screen_size.height() / base_height
        scaling = max(0.8, min(scale_w, scale_h, 3.0))
        return scaling
    except Exception:
        return 1.0

# === AI Latin Name Fetch Logic ===
def get_latin_name_wikipedia(common_name):
    try:
        page = wikipedia.page(common_name + " homeopathy", auto_suggest=True)
        summary = page.content.split('\n')[0]
        if '(' in summary:
            latin_candidate = summary.split('(')[1].split(')')[0]
            if len(latin_candidate.split()) > 1 and all(x.isalpha() or x.isspace() for x in latin_candidate):
                return latin_candidate
        return None
    except Exception:
        return None

def get_latin_name_clinicwala(common_name):
    url = "https://www.clinicwala.com/cure_homeopathy-homeopathy-namesofhomeopathicmedicines.html"
    try:
        r = requests.get(url, timeout=10)
        soup = BeautifulSoup(r.text, 'html.parser')
        rows = soup.select('table')[0].find_all('tr')
        for row in rows:
            cells = row.find_all('td')
            if len(cells) >= 2:
                common = cells[0].text.strip().lower()
                latin = cells[1].text.strip()
                if common == common_name.strip().lower():
                    return latin
    except Exception:
        pass
    return None

def fetch_latin_name(common_name):
    latin = get_latin_name_wikipedia(common_name)
    if latin:
        return latin
    latin = get_latin_name_clinicwala(common_name)
    if latin:
        return latin
    return "No online Latin name found."

class HomeoLabelApp(QtWidgets.QWidget):
    def __init__(self, scaling=1.0):
        super().__init__()
        self.setWindowTitle("🏥 Homeopathy Label Generator")
        self.setWindowFlags(QtCore.Qt.Window)
        self.showMaximized()
        self.scaling = scaling if scaling and scaling > 0 else 1.0
        self.font_size_med = int(8 * self.scaling)
        self.top_offset = 6.0 * self.scaling
        self.records_folder = "records"
        os.makedirs(self.records_folder, exist_ok=True)
        logging.basicConfig(filename="records/error_log.txt", level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
        self.excel_file = os.path.join(self.records_folder, 'records.xlsx')
        self.autocomplete_file = os.path.join(self.records_folder, 'autocomplete.json')
        self.remedies_file = 'remedies.xlsx'
        self.df_remedies = None
        self.load_remedies()
        self.autocomplete_data = self.load_autocomplete()
        self.record_buffer = []
        self.init_ui()

    def load_remedies(self):
        if not os.path.exists(self.remedies_file):
            df = pd.DataFrame({
                'latin_col': ['Arnica montana', 'Bryonia alba', 'Atropa belladonna'],
                'common_col': ['Arnica', 'Bryonia', 'Belladonna']
            })
            df.to_excel(self.remedies_file, index=False, engine="openpyxl")
        try:
            self.df_remedies = pd.read_excel(self.remedies_file, engine="openpyxl")
            self.df_remedies.fillna('', inplace=True)
            logging.info("Remedies loaded successfully.")
        except Exception as e:
            logging.error(f"Failed to load remedies.xlsx: {e}")
            QMessageBox.critical(self, "Error", f"Failed to load remedies.xlsx:\n{e}")

    def load_autocomplete(self):
        if os.path.exists(self.autocomplete_file):
            try:
                with open(self.autocomplete_file, "r") as f:
                    return json.load(f)
            except Exception as e:
                logging.warning(f"Autocomplete file load failed: {e}")
                return {}
        return {}

    def save_autocomplete(self):
        backup_file = self.autocomplete_file.replace(".json", "_backup.json")
        try:
            if Path(self.autocomplete_file).exists():
                os.replace(self.autocomplete_file, backup_file)
            tmp_file = self.autocomplete_file + ".tmp"
            with open(tmp_file, "w") as f:
                json.dump(self.autocomplete_data, f)
            os.replace(tmp_file, self.autocomplete_file)
            logging.info("Autocomplete saved successfully.")
        except Exception as e:
            logging.error(f"Failed to save autocomplete.json: {e}")

    def init_ui(self):
        font_medsz = int(18 * self.scaling)
        font_lbl = int(16 * self.scaling)
        pad_box = int(10 * self.scaling)
        pad_btn = int(12 * self.scaling)
        main_layout = QtWidgets.QHBoxLayout(self)
        main_layout.setSpacing(int(30 * self.scaling))
        left_panel = QtWidgets.QVBoxLayout()
        left_panel.setSpacing(int(20 * self.scaling))
        self.medicine_search = QtWidgets.QLineEdit()
        self.medicine_search.setPlaceholderText("Type medicine name (Latin or Common)")
        self.medicine_search.setStyleSheet(f"font-size:{font_medsz}pt; padding:{pad_box}px;")
        self.medicine_search.textChanged.connect(self.update_suggestions)
        left_panel.addWidget(QtWidgets.QLabel("Find Medicine", styleSheet=f"font-weight:bold; font-size:{font_lbl}pt;"))
        left_panel.addWidget(self.medicine_search)

        # == Add AI Fetch Latin Name Button ==
        self.fetch_latin_btn = QtWidgets.QPushButton("Fetch Latin Name (AI)")
        self.fetch_latin_btn.setStyleSheet(f"font-size:{int(14*self.scaling)}pt; margin-bottom:8px;")
        self.fetch_latin_btn.clicked.connect(self.on_fetch_latin_name)
        left_panel.addWidget(self.fetch_latin_btn)

        self.suggestion_table = QtWidgets.QTableWidget()
        self.suggestion_table.setColumnCount(2)
        self.suggestion_table.setHorizontalHeaderLabels(["Common Name", "Latin Name"])
        self.suggestion_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.suggestion_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.suggestion_table.setStyleSheet(f"font-size:{int(14 * self.scaling)}pt;")
        col_width_common = int(150 * self.scaling)
        col_width_latin  = int(200 * self.scaling)
        self.suggestion_table.setColumnWidth(0, col_width_common)
        self.suggestion_table.setColumnWidth(1, col_width_latin)
        header = self.suggestion_table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.Fixed)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.Fixed)
        self.suggestion_table.cellClicked.connect(self.on_suggestion_clicked)
        left_panel.addWidget(QtWidgets.QLabel("Suggestions", styleSheet=f"font-weight:bold; font-size:{int(15 * self.scaling)}pt;"))
        left_panel.addWidget(self.suggestion_table)
        self.add_new_btn = QtWidgets.QPushButton("Add New Medicine")
        self.add_new_btn.setStyleSheet(f"padding:{pad_box}px; font-size:{int(14*self.scaling)}pt; font-weight:bold;")
        self.add_new_btn.setToolTip("Add a medicine not in the list")
        self.add_new_btn.clicked.connect(self.add_new_medicine)
        left_panel.addWidget(self.add_new_btn)
        left_panel.addStretch()
        main_layout.addLayout(left_panel, 2)
        right_panel = QtWidgets.QVBoxLayout()
        right_panel.setSpacing(int(25 * self.scaling))
        self.selected_medicine_label = QtWidgets.QLabel("MEDICINE: ")
        self.selected_medicine_label.setStyleSheet(f"font-weight:bold; font-size:{int(24*self.scaling)}pt; margin-bottom:{int(16*self.scaling)}px;")
        right_panel.addWidget(self.selected_medicine_label)
        form = QtWidgets.QFormLayout()
        form.setVerticalSpacing(int(20 * self.scaling))
        self.potency_input = QtWidgets.QComboBox()
        form.addRow(QtWidgets.QLabel("Potency:", styleSheet=f"font-size:{int(15*self.scaling)}pt; font-weight:bold;"), self.potency_input)
        self.potency_input.setEditable(True)
        self.potency_input.setStyleSheet(f"font-size:{int(16*self.scaling)}pt; padding:{int(7*self.scaling)}px;")
        pot_list = self.autocomplete_data.get("potency", [])
        self.potency_input.addItems(pot_list)
        self.potency_input.setCompleter(QCompleter(pot_list))
        self.potency_input.currentTextChanged.connect(self.update_preview)

        self.dose_input = QtWidgets.QComboBox()
        form.addRow(QtWidgets.QLabel("Dose:", styleSheet=f"font-size:{int(15*self.scaling)}pt; font-weight:bold;"), self.dose_input)
        self.dose_input.setEditable(True)
        self.dose_input.setStyleSheet(f"font-size:{int(16*self.scaling)}pt; padding:{int(7*self.scaling)}px;")
        dose_list = self.autocomplete_data.get("dose", [])
        self.dose_input.addItems(dose_list)
        self.dose_input.setCompleter(QCompleter(dose_list))
        self.dose_input.currentTextChanged.connect(self.update_preview)

        self.time_input = QtWidgets.QComboBox()
        form.addRow(QtWidgets.QLabel("Time:", styleSheet=f"font-size:{int(15*self.scaling)}pt; font-weight:bold;"), self.time_input)
        self.time_input.setEditable(True)
        self.time_input.setStyleSheet(f"font-size:{int(16*self.scaling)}pt; padding:{int(7*self.scaling)}px;")
        time_list = self.autocomplete_data.get("time", [])
        self.time_input.addItems(time_list)
        self.time_input.setCompleter(QCompleter(time_list))
        self.time_input.currentTextChanged.connect(self.update_preview)

        self.shop_input = QtWidgets.QComboBox()
        form.addRow(QtWidgets.QLabel("Shop Name:", styleSheet=f"font-size:{int(15*self.scaling)}pt; font-weight:bold;"), self.shop_input)
        self.shop_input.setEditable(True)
        self.shop_input.setStyleSheet(f"font-size:{int(16*self.scaling)}pt; padding:{int(7*self.scaling)}px;")
        shop_list = self.autocomplete_data.get("shop", [])
        self.shop_input.addItems(shop_list)
        self.shop_input.setCompleter(QCompleter(shop_list))
        self.shop_input.currentTextChanged.connect(self.update_preview)

        self.branch_phone_input = QtWidgets.QComboBox()
        form.addRow(QtWidgets.QLabel("Branch/Phone:", styleSheet=f"font-size:{int(15*self.scaling)}pt; font-weight:bold;"), self.branch_phone_input)
        self.branch_phone_input.setEditable(True)
        self.branch_phone_input.setStyleSheet(f"font-size:{int(16*self.scaling)}pt; padding:{int(7*self.scaling)}px;")
        branch_list = self.autocomplete_data.get("branch", [])
        self.branch_phone_input.addItems(branch_list)
        self.branch_phone_input.setCompleter(QCompleter(branch_list))
        self.branch_phone_input.currentTextChanged.connect(self.update_preview)

        right_panel.addLayout(form)
        self.preview_frame = QtWidgets.QFrame()
        self.preview_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.preview_frame.setStyleSheet("background-color:#f8f8f8; border:2px solid #888;")
        preview_layout = QtWidgets.QVBoxLayout(self.preview_frame)
        self.preview_line1 = QtWidgets.QLabel("")
        self.preview_line1.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_line1.setStyleSheet(f"font-weight:bold; font-size:{int(28*self.scaling)}pt;")
        preview_layout.addWidget(self.preview_line1)
        self.preview_line2 = QtWidgets.QLabel("")
        self.preview_line2.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_line2.setStyleSheet(f"font-weight:bold; font-size:{int(24*self.scaling)}pt;")
        preview_layout.addWidget(self.preview_line2)
        self.preview_line3 = QtWidgets.QLabel("")
        self.preview_line3.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_line3.setStyleSheet(f"font-size:{int(20*self.scaling)}pt;")
        self.preview_line3.setWordWrap(True)
        preview_layout.addWidget(self.preview_line3)
        self.preview_line4 = QtWidgets.QLabel("")
        self.preview_line4.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_line4.setStyleSheet(f"font-size:{int(18*self.scaling)}pt;")
        preview_layout.addWidget(self.preview_line4)
        self.preview_line5 = QtWidgets.QLabel("")
        self.preview_line5.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_line5.setStyleSheet(f"font-size:{int(14*self.scaling)}pt;")
        preview_layout.addWidget(self.preview_line5)
        right_panel.addWidget(QtWidgets.QLabel("Label Preview", styleSheet=f"font-weight:bold;font-size:{int(18*self.scaling)}pt;"))
        right_panel.addWidget(self.preview_frame)
        controls_layout = QtWidgets.QHBoxLayout()
        self.font_size_spin = QtWidgets.QSpinBox()
        self.font_size_spin.setRange(8, 36)
        self.font_size_spin.setValue(self.font_size_med)
        self.font_size_spin.setStyleSheet(f"font-size:{int(16*self.scaling)}pt;")
        self.font_size_spin.valueChanged.connect(self.update_label_settings)
        controls_layout.addWidget(QtWidgets.QLabel("Font Size:", styleSheet=f"font-size:{int(16*self.scaling)}pt;"))
        controls_layout.addWidget(self.font_size_spin)
        self.printer_combo = QtWidgets.QComboBox()
        self.printer_combo.setStyleSheet(f"font-size:{int(15*self.scaling)}pt;")
        self.refresh_printers()
        controls_layout.addWidget(QtWidgets.QLabel("Printer:", styleSheet=f"font-size:{int(16*self.scaling)}pt;"))
        controls_layout.addWidget(self.printer_combo)
        self.printer_refresh_btn = QtWidgets.QPushButton("Refresh")
        self.printer_refresh_btn.setStyleSheet(f"padding:{pad_btn}px; font-size:{int(14*self.scaling)}pt; font-weight:bold;")
        self.printer_refresh_btn.clicked.connect(self.refresh_printers)
        controls_layout.addWidget(self.printer_refresh_btn)
        right_panel.addLayout(controls_layout)
        btn_layout = QtWidgets.QHBoxLayout()
        self.print_btn = QtWidgets.QPushButton("Preview PDF")
        self.print_btn.setStyleSheet(f'font-weight:bold; font-size:{int(18*self.scaling)}pt; padding:{int(14*self.scaling)}px;')
        self.print_btn.clicked.connect(self.print_label)
        btn_layout.addWidget(self.print_btn)
        self.direct_print_btn = QtWidgets.QPushButton("Direct Print")
        self.direct_print_btn.setStyleSheet(f'font-weight:bold; font-size:{int(18*self.scaling)}pt; padding:{int(14*self.scaling)}px;')
        self.direct_print_btn.clicked.connect(self.print_direct)
        btn_layout.addWidget(self.direct_print_btn)
        right_panel.addLayout(btn_layout)
        self.status = QtWidgets.QLabel("Ready")
        self.status.setStyleSheet(f"color: darkgreen; font-size:{int(16*self.scaling)}pt")
        right_panel.addWidget(self.status)
        right_panel.addStretch()
        main_layout.addLayout(right_panel, 3)

    def update_suggestions(self):
        text = self.medicine_search.text().lower().strip()
        self.suggestion_table.setRowCount(0)
        if not text:
            return
        for _, row in self.df_remedies.iterrows():
            common = str(row['common_col'])
            latin = str(row['latin_col'])
            if text in common.lower() or text in latin.lower():
                row_idx = self.suggestion_table.rowCount()
                self.suggestion_table.insertRow(row_idx)
                self.suggestion_table.setItem(row_idx, 0, QTableWidgetItem(common))
                self.suggestion_table.setItem(row_idx, 1, QTableWidgetItem(latin))

    def on_suggestion_clicked(self, row, column):
        item = self.suggestion_table.item(row, column)
        if item:
            self.medicine_search.setText(item.text())
            self.update_selected_medicine()

    def add_new_medicine(self):
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("Add New Medicine")
        layout = QtWidgets.QFormLayout(dialog)

        common_input = QtWidgets.QLineEdit()
        latin_input = QtWidgets.QLineEdit()

        # Optionally add an AI fetch button in dialog
        ai_btn = QtWidgets.QPushButton("Fetch Latin Name (AI)")
        def do_ai_fetch():
            c_name = common_input.text().strip()
            latin_result = fetch_latin_name(c_name)
            latin_input.setText(latin_result if latin_result and latin_result != "No online Latin name found." else "")
        ai_btn.clicked.connect(do_ai_fetch)

        layout.addRow("Common Name:", common_input)
        layout.addRow("Latin Name:", latin_input)
        layout.addRow("", ai_btn)

        buttons = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel,
            QtCore.Qt.Horizontal, dialog
        )
        layout.addRow(buttons)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)

        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            common = common_input.text().strip()
            latin = latin_input.text().strip()
            if not common or not latin:
                QMessageBox.warning(self, "Input Error", "Please enter both Common and Latin names.")
                return
            self.save_new_medicine(common, latin)
            self.medicine_search.setText(common)
            self.update_suggestions()

    def save_new_medicine(self, common_name, latin_name):
        exists = ((self.df_remedies['common_col'].str.lower() == common_name.lower()) |
                  (self.df_remedies['latin_col'].str.lower() == latin_name.lower())).any()
        if not exists:
            new_row = {'common_col': common_name, 'latin_col': latin_name}
            self.df_remedies = pd.concat([self.df_remedies, pd.DataFrame([new_row])], ignore_index=True)
            self.df_remedies.to_excel(self.remedies_file, index=False, engine='openpyxl')
            logging.info(f"New medicine added: {common_name} / {latin_name}")

    def update_selected_medicine(self):
        med_name = self.medicine_search.text().upper()
        self.selected_medicine_label.setText(f"MEDICINE: {med_name}")
        self.update_preview()

    def update_preview(self):
        med_name = self.medicine_search.text().strip().upper()
        potency = self.potency_input.currentText().upper()
        dose = self.dose_input.currentText()
        time_val = self.time_input.currentText()
        shop = self.shop_input.currentText().upper()
        branch_phone = self.branch_phone_input.currentText().upper()
        words = med_name.split()
        line1, line2 = "", ""
        for word in words:
            if len(line1 + " " + word) <= 18:
                line1 += (" " + word).strip()
            else:
                line2 += (" " + word).strip()
        self.preview_line1.setText(line1)
        self.preview_line2.setText(f"{line2} {potency}".strip())
        self.preview_line3.setText(f"{dose}   {time_val}")
        self.preview_line4.setText(f"{shop}")
        self.preview_line5.setText(f"{branch_phone}")

    def update_label_settings(self):
        self.font_size_med = self.font_size_spin.value()
        self.update_preview()
        logging.info(f"Font size changed to: {self.font_size_med}")

    def print_label(self):
        if not self.medicine_search.text().strip():
            QMessageBox.warning(self, "Missing Info", "Please enter or select a medicine before printing.")
            return
        try:
            pdf_file = os.path.join(self.records_folder, "label.pdf")
            self.generate_pdf(pdf_file)
            os.startfile(pdf_file)
            self.status.setText("Label preview opened and record saved.")
            logging.info(f"Label previewed for {self.medicine_search.text().strip()}")
        except Exception as e:
            logging.error(f"Print failed: {e}")
            QMessageBox.critical(self, "Error", f"Print failed: {e}")
            self.status.setText(f"Error: {e}")

    def print_direct(self):
        if not self.medicine_search.text().strip():
            QMessageBox.warning(self, "Missing Info", "Please enter or select a medicine before printing.")
            return
        self.refresh_printers()
        printer_name = self.printer_combo.currentText()
        if not printer_name:
            QMessageBox.warning(self, "No Printer", "Please select a printer first.")
            return
        if not self.check_printer_ready(printer_name):
            QMessageBox.critical(self, "Printer Error",
                f"Printer '{printer_name}' is not ready or not connected via USB.\n"
                "• Make sure printer is ON and connected\n"
                "• Try a different USB port/cable\n"
                "• Print a test page from Windows\n"
                "• Use 'Refresh Printers' in this app")
            return
        try:
            pdf_file = os.path.join(self.records_folder, "label.pdf")
            self.generate_pdf(pdf_file)
            result = win32api.ShellExecute(0, "printto", pdf_file, f'"{printer_name}"', ".", 0)
            if int(result) <= 32:
                raise OSError(f"ShellExecute error: {result}")
            self.status.setText(f"Label sent to {printer_name}.")
            logging.info(f"Label sent to printer: {printer_name}")
        except Exception as e:
            logging.error(f"Direct print failed: {e}")
            QMessageBox.critical(self, "Direct Print Failed",
                f"Printing failed: {e}\n\n"
                "Please check the printer connection, Windows status, and try manual print.")
            try:
                os.startfile(pdf_file)
                QMessageBox.information(self, "Manual Print",
                    "Direct print failed, but PDF is opened for manual printing.\n"
                    "Please print using your PDF viewer.")
                self.status.setText("PDF opened for manual print.")
            except Exception as e2:
                logging.error(f"Failed to open PDF for manual print: {e2}")
                self.status.setText("PDF manual print also failed.")

    def generate_pdf(self, pdf_file):
        med_name = self.medicine_search.text().strip().upper()
        potency = self.potency_input.currentText().upper()
        dose = self.dose_input.currentText()
        time_val = self.time_input.currentText()
        shop = self.shop_input.currentText().upper()
        branch_phone = self.branch_phone_input.currentText().upper()
        if not med_name:
            QMessageBox.warning(self, "Missing Info", "Please enter or select a medicine before printing.")
            return
        record = {"Medicine": med_name, "Potency": potency, "Dose": dose,
                  "Time": time_val, "Shop": shop, "Branch/Phone": branch_phone}
        self.record_buffer.append(record)
        if len(self.record_buffer) >= 10 or pdf_file:
            try:
                if os.path.exists(self.excel_file):
                    df_records = pd.read_excel(self.excel_file, engine="openpyxl")
                    df_records = pd.concat([df_records, pd.DataFrame(self.record_buffer)], ignore_index=True)
                else:
                    df_records = pd.DataFrame(self.record_buffer)
                df_records.to_excel(self.excel_file, index=False, engine="openpyxl")
                self.record_buffer.clear()
                logging.info("Records buffered and saved.")
            except PermissionError:
                QMessageBox.warning(self, "File Locked", "Please close 'records.xlsx' before saving again.")
                logging.warning("Excel file permission error encountered.")
                return
        for field, value in [("potency", potency), ("dose", dose), ("time", time_val),
                             ("shop", shop), ("branch", branch_phone)]:
            if value:
                lst = self.autocomplete_data.setdefault(field, [])
                if value not in lst:
                    lst.append(value)
        self.save_autocomplete()
        width_mm, height_mm = 50, 30
        c = canvas.Canvas(pdf_file, pagesize=(width_mm * mm, height_mm * mm))
        c.setLineWidth(1)
        c.rect(2 * mm, 2 * mm, (width_mm - 4) * mm, (height_mm - 4) * mm)
        y = height_mm * mm - self.top_offset * mm
        c.setFont("Helvetica-Bold", self.font_size_med)
        words = med_name.split()
        line1, line2 = "", ""
        for word in words:
            if len(line1 + " " + word) <= 18:
                line1 += (" " + word).strip()
            else:
                line2 += (" " + word).strip()
        c.drawCentredString((width_mm / 2) * mm, y, line1)
        y -= 5 * mm
        c.drawCentredString((width_mm / 2) * mm, y, f"{line2} {potency}".strip())
        y -= 6 * mm
        dose = dose if dose else ''
        time_val = time_val if time_val else ''
        dose_width = stringWidth(dose, "Helvetica-Bold", 8)
        time_width = stringWidth(time_val, "Helvetica", 8)
        gap = 6  # points spacing
        total_width = dose_width + (gap if dose and time_val else 0) + time_width
        center_x = (width_mm / 2) * mm
        start_x = center_x - total_width / 2
        c.setFont("Helvetica-Bold", 8)
        c.drawString(start_x, y, dose)
        if time_val:
            c.setFont("Helvetica", 8)
            c.drawString(start_x + dose_width + gap, y, time_val)
        y -= 5 * mm
        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString((width_mm / 2) * mm, y, f"{shop}")
        y -= 4 * mm
        c.drawCentredString((width_mm / 2) * mm, y, f"{branch_phone}")
        c.save()

    def refresh_printers(self):
        self.printer_combo.clear()
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            self.printer_combo.addItems(printers)
            logging.info(f"Refreshed printer list: {printers}")
            QtWidgets.QApplication.processEvents()
        except Exception as e:
            logging.error(f"Failed to refresh printers: {e}")

    def check_printer_ready(self, printer_name):
        if not printer_name:
            return False
        try:
            printer_handle = win32print.OpenPrinter(printer_name)
            printer_info = win32print.GetPrinter(printer_handle, 2)
            win32print.ClosePrinter(printer_handle)
            status = printer_info['Status']
            if status == 0 and printer_info['Attributes'] & win32print.PRINTER_ATTRIBUTE_LOCAL:
                return True
            logging.warning(f"Printer '{printer_name}' not ready or not local: {status}")
            return False
        except Exception as e:
            logging.error(f"Printer check failed: {e}")
            return False

    # === AI Fetch Button Handler ===
    def on_fetch_latin_name(self):
        common_name = self.medicine_search.text().strip()
        if not common_name:
            QMessageBox.warning(self, "Input Error", "Please enter a common name first.")
            return
        self.status.setText("Fetching from internet...")
        QtWidgets.QApplication.processEvents()
        latin = fetch_latin_name(common_name)
        self.status.setText("Ready")
        if latin and latin != "No online Latin name found.":
            QMessageBox.information(self, "Latin Name Found", f"Common: {common_name}\nLatin: {latin}")
        else:
            QMessageBox.information(self, "Latin Name Not Found", f"No online Latin name found for: {common_name}")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    scaling = get_system_scaling()
    w = HomeoLabelApp(scaling)
    w.show()
    sys.exit(app.exec_())
