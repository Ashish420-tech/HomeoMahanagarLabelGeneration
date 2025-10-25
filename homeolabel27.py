import sys
import os
import json
import logging
import pandas as pd
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QCompleter, QTableWidgetItem, QMessageBox
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import win32api
import win32print
from pathlib import Path

class HomeoLabelApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🏥 Homeopathy Label Generator")
        self.resize(1000, 700)
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
        self.font_size_med = 8
        self.top_offset = 6.0 # mm, fixed
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
        layout = QtWidgets.QVBoxLayout(self)
        self.selected_medicine_label = QtWidgets.QLabel("MEDICINE: ")
        self.selected_medicine_label.setStyleSheet("font-weight:bold; font-size:12pt;")
        layout.addWidget(self.selected_medicine_label)
        search_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(search_layout)
        self.medicine_search = QtWidgets.QLineEdit()
        self.medicine_search.setPlaceholderText("Type medicine name (Latin or Common)")
        self.medicine_search.setMinimumWidth(300)
        self.medicine_search.textChanged.connect(self.update_suggestions)
        search_layout.addWidget(self.medicine_search)
        self.suggestion_table = QtWidgets.QTableWidget()
        self.suggestion_table.setColumnCount(2)
        self.suggestion_table.setHorizontalHeaderLabels(["Common Name", "Latin Name"])
        self.suggestion_table.setMinimumWidth(400)
        self.suggestion_table.setMinimumHeight(200)
        self.suggestion_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.suggestion_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.suggestion_table.cellClicked.connect(self.on_suggestion_clicked)
        search_layout.addWidget(self.suggestion_table)
        self.add_new_btn = QtWidgets.QPushButton("Add New Medicine")
        self.add_new_btn.clicked.connect(self.add_new_medicine)
        layout.addWidget(self.add_new_btn)
        form = QtWidgets.QFormLayout()
        layout.addLayout(form)
        self.potency_input = QtWidgets.QComboBox()
        self.potency_input.setEditable(True)
        pot_list = self.autocomplete_data.get("potency", [])
        self.potency_input.addItems(pot_list)
        self.potency_input.setCompleter(QCompleter(pot_list))
        self.potency_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Potency:", self.potency_input)
        self.dose_input = QtWidgets.QComboBox()
        self.dose_input.setEditable(True)
        dose_list = self.autocomplete_data.get("dose", [])
        self.dose_input.addItems(dose_list)
        self.dose_input.setCompleter(QCompleter(dose_list))
        self.dose_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Dose:", self.dose_input)
        self.time_input = QtWidgets.QComboBox()
        self.time_input.setEditable(True)
        time_list = self.autocomplete_data.get("time", [])
        self.time_input.addItems(time_list)
        self.time_input.setCompleter(QCompleter(time_list))
        self.time_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Time:", self.time_input)
        self.shop_input = QtWidgets.QComboBox()
        self.shop_input.setEditable(True)
        shop_list = self.autocomplete_data.get("shop", [])
        self.shop_input.addItems(shop_list)
        self.shop_input.setCompleter(QCompleter(shop_list))
        self.shop_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Shop Name:", self.shop_input)
        self.branch_phone_input = QtWidgets.QComboBox()
        self.branch_phone_input.setEditable(True)
        branch_list = self.autocomplete_data.get("branch", [])
        self.branch_phone_input.addItems(branch_list)
        self.branch_phone_input.setCompleter(QCompleter(branch_list))
        self.branch_phone_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Branch/Phone:", self.branch_phone_input)
        self.preview_frame = QtWidgets.QFrame()
        self.preview_frame.setFrameShape(QtWidgets.QFrame.Box)
        self.preview_frame.setFixedSize(400, 180)
        preview_layout = QtWidgets.QVBoxLayout(self.preview_frame)
        self.preview_line1 = QtWidgets.QLabel("")
        self.preview_line1.setStyleSheet("font-weight:bold; font-size:10pt;")
        self.preview_line1.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line1)
        self.preview_line2 = QtWidgets.QLabel("")
        self.preview_line2.setStyleSheet("font-weight:bold; font-size:10pt;")
        self.preview_line2.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line2)
        self.preview_line3 = QtWidgets.QLabel("")
        self.preview_line3.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_line3.setWordWrap(True)
        preview_layout.addWidget(self.preview_line3)
        self.preview_line4 = QtWidgets.QLabel("")
        self.preview_line4.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line4)
        self.preview_line5 = QtWidgets.QLabel("")
        self.preview_line5.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line5)
        settings_group = QtWidgets.QGroupBox("Label Settings")
        settings_layout = QtWidgets.QHBoxLayout()
        settings_group.setLayout(settings_layout)
        self.font_size_spin = QtWidgets.QSpinBox()
        self.font_size_spin.setRange(6, 20)
        self.font_size_spin.setValue(self.font_size_med)
        self.font_size_spin.valueChanged.connect(self.update_label_settings)
        settings_layout.addWidget(QtWidgets.QLabel("Font Size:"))
        settings_layout.addWidget(self.font_size_spin)
        layout.addWidget(settings_group)
        printers_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(printers_layout)
        printers_layout.addWidget(QtWidgets.QLabel("Select Printer:"))
        self.printer_combo = QtWidgets.QComboBox()
        self.refresh_printers()
        printers_layout.addWidget(self.printer_combo)
        self.printer_refresh_btn = QtWidgets.QPushButton("Refresh Printers")
        self.printer_refresh_btn.clicked.connect(self.refresh_printers)
        printers_layout.addWidget(self.printer_refresh_btn)
        preview_btn_layout = QtWidgets.QHBoxLayout()
        preview_btn_layout.addWidget(self.preview_frame)
        self.print_btn = QtWidgets.QPushButton("Preview PDF")
        self.print_btn.clicked.connect(self.print_label)
        preview_btn_layout.addWidget(self.print_btn)
        self.direct_print_btn = QtWidgets.QPushButton("Direct Print")
        self.direct_print_btn.clicked.connect(self.print_direct)
        preview_btn_layout.addWidget(self.direct_print_btn)
        layout.addLayout(preview_btn_layout)
        self.status = QtWidgets.QLabel("Ready")
        layout.addWidget(self.status)

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

    def update_label_settings(self):
        self.font_size_med = self.font_size_spin.value()
        self.update_preview()
        logging.info(f"Font size changed to: {self.font_size_med}")

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
        self.suggestion_table.resizeColumnsToContents()

    def on_suggestion_clicked(self, row, column):
        item = self.suggestion_table.item(row, column)
        if item:
            self.medicine_search.setText(item.text())
            self.update_selected_medicine()

    def add_new_medicine(self):
        new_name, ok = QtWidgets.QInputDialog.getText(self, "Add New Medicine", "Enter medicine name:")
        if ok and new_name.strip():
            self.save_new_medicine(new_name.strip())
            self.medicine_search.setText(new_name.strip())
            self.update_suggestions()

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

    def save_new_medicine(self, med_name):
        exists = ((self.df_remedies['common_col'].str.lower() == med_name.lower()) |
                  (self.df_remedies['latin_col'].str.lower() == med_name.lower())).any()
        if not exists:
            new_row = {'common_col': med_name, 'latin_col': med_name}
            self.df_remedies = pd.concat([self.df_remedies, pd.DataFrame([new_row])], ignore_index=True)
            self.df_remedies.to_excel(self.remedies_file, index=False, engine='openpyxl')
            logging.info(f"New medicine added: {med_name}")

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
                # Windows ShellExecute errors are <= 32
                raise OSError(f"ShellExecute error: {result}")

            self.status.setText(f"Label sent to {printer_name}.")
            logging.info(f"Label sent to printer: {printer_name}")
        except Exception as e:
            logging.error(f"Direct print failed: {e}")
            QMessageBox.critical(self, "Direct Print Failed",
                f"Printing failed: {e}\n\n"
                "Please check the printer connection, Windows status, and try manual print.")
            # Automatic fallback: Open PDF for manual print
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
        y = height_mm * mm - self.top_offset * mm  # Fixed top offset always used
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
        c.setFont("Helvetica", 8)
        c.drawCentredString((width_mm / 2) * mm, y, f"{dose}   {time_val}")
        y -= 5 * mm
        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString((width_mm / 2) * mm, y, f"{shop}")
        y -= 4 * mm
        c.drawCentredString((width_mm / 2) * mm, y, f"{branch_phone}")
        c.save()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = HomeoLabelApp()
    w.show()
    sys.exit(app.exec_())
