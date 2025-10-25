import sys
import os
import json
import logging
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QCompleter
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch

# ------------------ Logging ------------------
logging.basicConfig(filename="error_log.txt", level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# ------------------ Homeopathy Label App ------------------
class HomeoLabelApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🏥 Homeopathy Label Generator")
        self.resize(550, 600)

        # Paths
        self.excel_file = 'remedies.xlsx'
        self.records_file = 'records.xlsx'
        self.autocomplete_file = 'autocomplete.json'

        # Load remedies
        self.df_remedies = None
        self.load_remedies()

        # Autocomplete data
        self.autocomplete_data = self.load_autocomplete()

        # UI
        self.init_ui()

    # ------------------ Load Remedies ------------------
    def load_remedies(self):
        if not os.path.exists(self.excel_file):
            df = pd.DataFrame({
                'latin_col': ['Arnica montana', 'Bryonia alba', 'Atropa belladonna'],
                'common_col': ['Arnica', 'Bryonia', 'Belladonna']
            })
            df.to_excel(self.excel_file, index=False, engine="openpyxl")
        try:
            self.df_remedies = pd.read_excel(self.excel_file, engine="openpyxl")
            self.df_remedies.fillna('', inplace=True)
        except Exception as e:
            logging.error("Failed to load remedies.xlsx: %s", str(e))
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load remedies.xlsx:\n{e}")

    # ------------------ Autocomplete ------------------
    def load_autocomplete(self):
        if os.path.exists(self.autocomplete_file):
            try:
                return json.load(open(self.autocomplete_file))
            except:
                return {}
        return {}

    # ------------------ UI Setup ------------------
    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # ----- Top: Selected Medicine Display -----
        self.selected_medicine_label = QtWidgets.QLabel("MEDICINE: ")
        self.selected_medicine_label.setStyleSheet("font-weight:bold; font-size:10pt;")
        layout.addWidget(self.selected_medicine_label)

        # ----- Medicine Search with Strict Autocomplete -----
        form = QtWidgets.QFormLayout()
        layout.addLayout(form)

        self.medicine_search = QtWidgets.QLineEdit()
        self.medicine_search.setPlaceholderText("Select medicine from list")
        self.medicine_search.textChanged.connect(self.update_medicine_suggestions)

        # Build completer from Excel
        medicine_list = list(dict.fromkeys(
            list(self.df_remedies['common_col']) + list(self.df_remedies['latin_col'])
        ))
        self.completer = QCompleter(medicine_list)
        self.completer.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.medicine_search.setCompleter(self.completer)
        self.completer.activated.connect(self.medicine_selected)

        form.addRow("Medicine Search:", self.medicine_search)

        # Potency dropdown
        self.potency_input = QtWidgets.QComboBox()
        self.potency_input.setEditable(True)
        pot_list = self.autocomplete_data.get("potency", [])
        self.potency_input.addItems(pot_list)
        self.potency_input.setCompleter(QCompleter(pot_list))
        self.potency_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Potency:", self.potency_input)

        # Dose
        self.dose_input = QtWidgets.QSpinBox()
        self.dose_input.setRange(1, 99)
        self.dose_input.valueChanged.connect(self.update_preview)
        form.addRow("Dose:", self.dose_input)

        # Quantity
        self.quantity_input = QtWidgets.QLineEdit()
        self.quantity_input.setPlaceholderText("Enter quantity")
        self.quantity_input.textChanged.connect(self.update_preview)
        form.addRow("Quantity:", self.quantity_input)

        # Time
        self.time_input = QtWidgets.QLineEdit()
        self.time_input.setPlaceholderText("Enter times like 8-12-4")
        self.time_input.textChanged.connect(self.update_preview)
        form.addRow("Time:", self.time_input)

        # Shop Name
        self.shop_input = QtWidgets.QLineEdit()
        self.shop_input.setPlaceholderText("Enter Shop Name, LIC No, Address")
        self.shop_input.textChanged.connect(self.update_preview)
        form.addRow("Shop Name:", self.shop_input)

        # ----- Label Preview at Bottom -----
        self.preview_frame = QtWidgets.QFrame()
        self.preview_frame.setFrameShape(QtWidgets.QFrame.Box)
        self.preview_frame.setFixedSize(400, 180)
        preview_layout = QtWidgets.QVBoxLayout(self.preview_frame)

        self.preview_line1 = QtWidgets.QLabel("")  # Medicine + Potency
        self.preview_line1.setStyleSheet("font-weight:bold; font-size:8pt;")
        preview_layout.addWidget(self.preview_line1)

        self.preview_line2 = QtWidgets.QLabel("")  # Dose + Quantity
        preview_layout.addWidget(self.preview_line2)

        self.preview_time = QtWidgets.QLabel("")  # Time
        preview_layout.addWidget(self.preview_time)

        self.preview_shop = QtWidgets.QLabel("")  # Shop
        self.preview_shop.setStyleSheet("font-weight:bold;")
        preview_layout.addWidget(self.preview_shop)

        layout.addWidget(self.preview_frame)

        # ----- Print Button -----
        self.print_btn = QtWidgets.QPushButton("Generate Label")
        self.print_btn.clicked.connect(self.print_label)
        layout.addWidget(self.print_btn)

        # Status
        self.status = QtWidgets.QLabel("Ready")
        layout.addWidget(self.status)

    # ------------------ Medicine Selection ------------------
    def update_medicine_suggestions(self):
        # Only suggestions from Excel
        text = self.medicine_search.text().lower()
        suggestions = []
        for _, row in self.df_remedies.iterrows():
            if text in str(row['common_col']).lower() or text in str(row['latin_col']).lower():
                suggestions.append(row['common_col'])
        self.completer.model().setStringList(list(dict.fromkeys(suggestions)))

    def medicine_selected(self, text):
        # User selected medicine from list
        self.medicine_search.setText(text)
        self.update_selected_medicine()

    def update_selected_medicine(self):
        med_name = self.medicine_search.text().upper()
        self.selected_medicine_label.setText(f"MEDICINE: {med_name}")
        self.update_preview()

    # ------------------ Update Preview ------------------
    def update_preview(self):
        med_name = self.medicine_search.text().upper()
        potency = self.potency_input.currentText().upper()
        dose = self.dose_input.value()
        quantity = self.quantity_input.text()
        time_val = self.time_input.text()
        shop = self.shop_input.text().upper()

        self.preview_line1.setText(f"{med_name} {potency}")
        self.preview_line2.setText(f"Dose: {dose}  Qty: {quantity}")
        self.preview_time.setText(f"Time: {time_val}")
        self.preview_shop.setText(f"SHOP: {shop}")

    # ------------------ Generate PDF Label ------------------
    def print_label(self):
        try:
            med_name = self.medicine_search.text().upper()
            if med_name not in list(self.df_remedies['common_col']):
                QtWidgets.QMessageBox.warning(self, "Error", "Select a valid medicine from Excel list")
                return

            potency = self.potency_input.currentText().upper()
            dose = str(self.dose_input.value())
            quantity = self.quantity_input.text()
            time_val = self.time_input.text()
            shop = self.shop_input.text().upper()

            # ------------------ Save Autocomplete ------------------
            for field, value in [("medicine", med_name), ("potency", potency),
                                 ("quantity", quantity), ("shop", shop)]:
                if value:
                    lst = self.autocomplete_data.setdefault(field, [])
                    if value not in lst:
                        lst.append(value)
            with open(self.autocomplete_file, "w") as f:
                json.dump(self.autocomplete_data, f)

            # ------------------ Save Record ------------------
            record = {
                "Medicine": med_name,
                "Potency": potency,
                "Dose": dose,
                "Quantity": quantity,
                "Time": time_val,
                "Shop": shop
            }
            if os.path.exists(self.records_file):
                df_records = pd.read_excel(self.records_file, engine="openpyxl")
                df_records = pd.concat([df_records, pd.DataFrame([record])], ignore_index=True)
            else:
                df_records = pd.DataFrame([record])
            df_records.to_excel(self.records_file, index=False, engine="openpyxl")

            # ------------------ Generate PDF ------------------
            pdf_file = "label.pdf"
            c = canvas.Canvas(pdf_file, pagesize=(2*inch, 1*inch))

            # Draw border
            c.setLineWidth(0.5)
            c.rect(0.05*inch, 0.05*inch, 1.9*inch, 0.9*inch)

            # Label content
            y = 0.85*inch
            c.setFont("Helvetica-Bold", 8)
            c.drawString(0.1*inch, y, f"{med_name} {potency}")

            y -= 0.15*inch
            c.setFont("Helvetica", 8)
            c.drawString(0.1*inch, y, f"Dose: {dose}  Qty: {quantity}")

            y -= 0.15*inch
            c.drawString(0.1*inch, y, f"Time: {time_val}")

            y -= 0.15*inch
            c.setFont("Helvetica-Bold", 8)
            c.drawString(0.1*inch, y, f"SHOP: {shop}")

            c.save()

            # Open PDF for printing
            if sys.platform == "win32":
                os.startfile(pdf_file)
            else:
                os.system(f"open {pdf_file}")

            self.status.setText("Label created, opened, and record saved.")

        except Exception as e:
            logging.error("Print/Save failed: %s", str(e))
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed: {e}")
            self.status.setText(f"Error: {e}")


# ------------------ Run App ------------------
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = HomeoLabelApp()
    w.show()
    sys.exit(app.exec_())
