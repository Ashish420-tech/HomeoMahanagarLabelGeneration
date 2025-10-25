import sys
import os
import json
import logging
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QCompleter
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors

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

        # ----- Medicine Search -----
        form = QtWidgets.QFormLayout()
        layout.addLayout(form)

        self.medicine_search = QtWidgets.QLineEdit()
        self.medicine_search.setPlaceholderText("Search medicine name...")
        self.medicine_search.textChanged.connect(self.update_medicine_suggestions)
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

        # Branch / Phone
        self.branch_phone_input = QtWidgets.QLineEdit()
        self.branch_phone_input.setPlaceholderText("Enter Branch and Mobile like SHAPOORJI☎8240297412")
        self.branch_phone_input.textChanged.connect(self.update_preview)
        form.addRow("Branch / Phone:", self.branch_phone_input)

        # ----- Label Preview at Bottom -----
        self.preview_frame = QtWidgets.QFrame()
        self.preview_frame.setFrameShape(QtWidgets.QFrame.Box)
        self.preview_frame.setFixedSize(400, 180)
        preview_layout = QtWidgets.QVBoxLayout(self.preview_frame)

        self.preview_line1 = QtWidgets.QLabel("")  # Medicine + Potency
        self.preview_line1.setStyleSheet("font-weight:bold; font-size:8pt;")
        preview_layout.addWidget(self.preview_line1)

        self.preview_line2 = QtWidgets.QLabel("")  # Quantity
        preview_layout.addWidget(self.preview_line2)

        self.preview_line3 = QtWidgets.QLabel("")  # Dose + Time
        preview_layout.addWidget(self.preview_line3)

        self.preview_line4 = QtWidgets.QLabel("")  # Shop + Address
        self.preview_line4.setStyleSheet("font-weight:bold;")
        preview_layout.addWidget(self.preview_line4)

        self.preview_line5 = QtWidgets.QLabel("")  # Branch + Phone
        preview_layout.addWidget(self.preview_line5)

        layout.addWidget(self.preview_frame)

        # ----- Print Button -----
        self.print_btn = QtWidgets.QPushButton("Generate Label")
        self.print_btn.clicked.connect(self.print_label)
        layout.addWidget(self.print_btn)

        # Status
        self.status = QtWidgets.QLabel("Ready")
        layout.addWidget(self.status)

    # ------------------ Medicine Search Suggestions ------------------
    def update_medicine_suggestions(self):
        text = self.medicine_search.text().lower()
        suggestions = []
        for _, row in self.df_remedies.iterrows():
            if text in str(row['common_col']).lower() or text in str(row['latin_col']).lower():
                suggestions.append(row['common_col'])
        if suggestions:
            completer = QCompleter(list(dict.fromkeys(suggestions)))
            self.medicine_search.setCompleter(completer)
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
        branch_phone = self.branch_phone_input.text().upper()

        self.preview_line1.setText(f"{med_name} {potency}")
        self.preview_line2.setText(f"{quantity}")
        self.preview_line3.setText(f"{dose} pills   {time_val}")
        self.preview_line4.setText(f"{shop}")
        self.preview_line5.setText(f"{branch_phone}")

    # ------------------ Generate PDF Label ------------------
    def print_label(self):
        try:
            med_name = self.medicine_search.text().upper()
            potency = self.potency_input.currentText().upper()
            dose = str(self.dose_input.value())
            quantity = self.quantity_input.text()
            time_val = self.time_input.text()
            shop = self.shop_input.text().upper()
            branch_phone = self.branch_phone_input.text().upper()

            if not med_name:
                QtWidgets.QMessageBox.warning(self, "Error", "Medicine name is required")
                return

            # Save autocomplete
            for field, value in [("medicine", med_name), ("potency", potency),
                                 ("quantity", quantity), ("shop", shop)]:
                if value:
                    lst = self.autocomplete_data.setdefault(field, [])
                    if value not in lst:
                        lst.append(value)
            with open(self.autocomplete_file, "w") as f:
                json.dump(self.autocomplete_data, f)

            # Save record
            record = {
                "Medicine": med_name,
                "Potency": potency,
                "Dose": dose,
                "Quantity": quantity,
                "Time": time_val,
                "Shop": shop,
                "BranchPhone": branch_phone
            }
            if os.path.exists(self.records_file):
                df_records = pd.read_excel(self.records_file, engine="openpyxl")
                df_records = pd.concat([df_records, pd.DataFrame([record])], ignore_index=True)
            else:
                df_records = pd.DataFrame([record])
            df_records.to_excel(self.records_file, index=False, engine="openpyxl")

            # ------------------ Generate PDF ------------------
            pdf_file = "label_45x25.pdf"
            label_width = 45 * mm
            label_height = 25 * mm
            c = canvas.Canvas(pdf_file, pagesize=(label_width, label_height))

            # Solid bold border
            c.setLineWidth(1)
            c.setStrokeColor(colors.black)
            c.rect(1*mm, 1*mm, label_width - 2*mm, label_height - 2*mm)

            def draw_center(text, y, font="Helvetica", size=8, bold=False):
                c.setFillColor(colors.black)
                if bold:
                    c.setFont("Helvetica-Bold", size)
                else:
                    c.setFont(font, size)
                text_width = c.stringWidth(text, font, size)
                x = (label_width - text_width)/2
                c.drawString(x, y, text)

            y = label_height - 4*mm
            draw_center(f"{med_name} {potency}", y, size=9, bold=True)
            y -= 5*mm
            draw_center(f"{quantity}", y, size=8)
            y -= 5*mm
            draw_center(f"{dose} pills   {time_val}", y, size=8)
            y -= 5*mm
            draw_center(f"{shop}", y, size=7, bold=True)
            y -= 4*mm
            draw_center(f"{branch_phone}", y, size=7)

            c.save()

            # Open PDF for preview
            if sys.platform == "win32":
                os.startfile(pdf_file)
            elif sys.platform == "darwin":
                os.system(f"open '{pdf_file}'")
            else:
                os.system(f"xdg-open '{pdf_file}'")

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
