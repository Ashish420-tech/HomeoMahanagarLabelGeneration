import sys
import os
import json
import logging
import datetime
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QCompleter, QTableWidgetItem, QSlider
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch

# ----------------- Central records folder -----------------
records_folder = "records"
if not os.path.exists(records_folder):
    os.makedirs(records_folder)

# File paths
records_file = os.path.join(records_folder, "records.xlsx")  # single Excel file
autocomplete_file = os.path.join(records_folder, "autocomplete.json")
error_log_file = os.path.join(records_folder, "error_log.txt")

# Configure logging
logging.basicConfig(filename=error_log_file, level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')


# ----------------- Homeopathy Label Generator -----------------
class HomeoLabelApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🏥 Homeopathy Label Generator")
        self.resize(950, 700)

        # Load remedies
        self.remedies_file = "remedies.xlsx"
        self.df_remedies = None
        self.load_remedies()

        # Load autocomplete
        self.autocomplete_data = self.load_autocomplete()

        # Border positions
        self.border_x = 0.05  # left margin in inch
        self.border_y = 0.05  # bottom margin in inch
        self.border_width = 1.67  # width in inch
        self.border_height = 0.88  # height in inch
        self.top_offset = 0.88  # starting top position for text

        # Initialize UI
        self.init_ui()

    # ---------------- Load remedies ----------------
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
        except Exception as e:
            logging.error(f"Failed to load remedies.xlsx: {e}")
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load remedies.xlsx:\n{e}")

    # ---------------- Load autocomplete ----------------
    def load_autocomplete(self):
        if os.path.exists(autocomplete_file):
            try:
                return json.load(open(autocomplete_file))
            except:
                return {}
        return {}

    # ---------------- Initialize UI ----------------
    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # Top: Selected medicine
        self.selected_medicine_label = QtWidgets.QLabel("MEDICINE: ")
        self.selected_medicine_label.setStyleSheet("font-weight:bold; font-size:10pt;")
        layout.addWidget(self.selected_medicine_label)

        # Search + Suggestions
        search_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(search_layout)

        # Medicine search input
        self.medicine_search = QtWidgets.QLineEdit()
        self.medicine_search.setPlaceholderText("Type medicine name (Latin or Common)")
        self.medicine_search.textChanged.connect(self.update_suggestions)
        search_layout.addWidget(self.medicine_search)

        # Suggestion Table
        self.suggestion_table = QtWidgets.QTableWidget()
        self.suggestion_table.setColumnCount(2)
        self.suggestion_table.setHorizontalHeaderLabels(["Common Name", "Latin Name"])
        self.suggestion_table.setMinimumWidth(400)
        self.suggestion_table.setMinimumHeight(200)
        self.suggestion_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.suggestion_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.suggestion_table.cellClicked.connect(self.on_suggestion_clicked)
        search_layout.addWidget(self.suggestion_table)

        # Add new medicine button
        self.add_new_btn = QtWidgets.QPushButton("Add New Medicine")
        self.add_new_btn.clicked.connect(self.add_new_medicine)
        layout.addWidget(self.add_new_btn)

        # Form inputs
        form = QtWidgets.QFormLayout()
        layout.addLayout(form)

        # Potency
        self.potency_input = QtWidgets.QComboBox()
        self.potency_input.setEditable(True)
        pot_list = self.autocomplete_data.get("potency", [])
        self.potency_input.addItems(pot_list)
        self.potency_input.setCompleter(QCompleter(pot_list))
        self.potency_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Potency:", self.potency_input)

        # Dose
        self.dose_input = QtWidgets.QLineEdit()
        self.dose_input.setPlaceholderText("Enter dose")
        self.dose_input.textChanged.connect(self.update_preview)
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

        # Shop
        self.shop_input = QtWidgets.QLineEdit()
        self.shop_input.setPlaceholderText("Enter Shop Name")
        self.shop_input.textChanged.connect(self.update_preview)
        form.addRow("Shop Name:", self.shop_input)

        # Branch/Phone
        self.branch_phone_input = QtWidgets.QLineEdit()
        self.branch_phone_input.setPlaceholderText("Enter Branch and Phone")
        self.branch_phone_input.textChanged.connect(self.update_preview)
        form.addRow("Branch/Phone:", self.branch_phone_input)

        # ---------------- Label Preview ----------------
        self.preview_frame = QtWidgets.QFrame()
        self.preview_frame.setFrameShape(QtWidgets.QFrame.Box)
        self.preview_frame.setFixedSize(400, 180)
        preview_layout = QtWidgets.QVBoxLayout(self.preview_frame)

        self.preview_line1 = QtWidgets.QLabel("")  # Medicine + Potency
        self.preview_line1.setStyleSheet("font-weight:bold; font-size:10pt;")
        self.preview_line1.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line1)

        self.preview_line2 = QtWidgets.QLabel("")  # Quantity
        self.preview_line2.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line2)

        self.preview_line3 = QtWidgets.QLabel("")  # Dose + Time
        self.preview_line3.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line3)

        self.preview_line4 = QtWidgets.QLabel("")  # Shop
        self.preview_line4.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line4)

        self.preview_line5 = QtWidgets.QLabel("")  # Branch/Phone
        self.preview_line5.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line5)

        layout.addWidget(self.preview_frame)

        # ---------------- Border adjustments sliders ----------------
        slider_layout = QtWidgets.QFormLayout()
        layout.addLayout(slider_layout)

        self.slider_top = QSlider(QtCore.Qt.Horizontal)
        self.slider_top.setMinimum(0)
        self.slider_top.setMaximum(200)
        self.slider_top.setValue(int(self.top_offset*100))
        self.slider_top.valueChanged.connect(self.update_border)
        slider_layout.addRow("Top offset (PDF):", self.slider_top)

        self.slider_width = QSlider(QtCore.Qt.Horizontal)
        self.slider_width.setMinimum(50)
        self.slider_width.setMaximum(200)
        self.slider_width.setValue(int(self.border_width*100))
        self.slider_width.valueChanged.connect(self.update_border)
        slider_layout.addRow("Label width (PDF):", self.slider_width)

        self.slider_height = QSlider(QtCore.Qt.Horizontal)
        self.slider_height.setMinimum(50)
        self.slider_height.setMaximum(200)
        self.slider_height.setValue(int(self.border_height*100))
        self.slider_height.valueChanged.connect(self.update_border)
        slider_layout.addRow("Label height (PDF):", self.slider_height)

        # Generate Label Button
        self.print_btn = QtWidgets.QPushButton("Generate Label")
        self.print_btn.clicked.connect(self.print_label)
        layout.addWidget(self.print_btn)

        # Status
        self.status = QtWidgets.QLabel("Ready")
        layout.addWidget(self.status)

    # ---------------- Update suggestions ----------------
    def update_suggestions(self):
        text = self.medicine_search.text().lower()
        self.suggestion_table.setRowCount(0)
        for _, row in self.df_remedies.iterrows():
            common = str(row['common_col'])
            latin = str(row['latin_col'])
            if text in common.lower() or text in latin.lower():
                row_idx = self.suggestion_table.rowCount()
                self.suggestion_table.insertRow(row_idx)
                self.suggestion_table.setItem(row_idx, 0, QTableWidgetItem(common))
                self.suggestion_table.setItem(row_idx, 1, QTableWidgetItem(latin))

    # ---------------- User selects name ----------------
    def on_suggestion_clicked(self, row, column):
        item = self.suggestion_table.item(row, column)
        if item:
            self.medicine_search.setText(item.text())
            self.update_selected_medicine()

    # ---------------- Add new medicine ----------------
    def add_new_medicine(self):
        new_name, ok = QtWidgets.QInputDialog.getText(self, "Add New Medicine", "Enter medicine name:")
        if ok and new_name.strip():
            self.save_new_medicine(new_name.strip())
            self.medicine_search.setText(new_name.strip())
            self.update_suggestions()

    # ---------------- Update selected medicine ----------------
    def update_selected_medicine(self):
        med_name = self.medicine_search.text().upper()
        self.selected_medicine_label.setText(f"MEDICINE: {med_name}")
        self.update_preview()

    # ---------------- Update preview ----------------
    def update_preview(self):
        med_name = self.medicine_search.text().upper()
        potency = self.potency_input.currentText().upper()
        dose = self.dose_input.text()
        quantity = self.quantity_input.text()
        time_val = self.time_input.text()
        shop = self.shop_input.text().upper()
        branch_phone = self.branch_phone_input.text().upper()

        self.preview_line1.setText(f"{med_name} {potency}")
        self.preview_line2.setText(f"{quantity}")
        self.preview_line3.setText(f"{dose}   {time_val}")  # pill removed
        self.preview_line4.setText(f"{shop}")
        self.preview_line5.setText(f"{branch_phone}")

    # ---------------- Save new medicine ----------------
    def save_new_medicine(self, med_name):
        med_name = med_name.strip()
        exists = ((self.df_remedies['common_col'].str.lower() == med_name.lower()) |
                  (self.df_remedies['latin_col'].str.lower() == med_name.lower())).any()
        if not exists:
            new_row = {'common_col': med_name, 'latin_col': med_name}
            self.df_remedies = pd.concat([self.df_remedies, pd.DataFrame([new_row])], ignore_index=True)
            self.df_remedies.to_excel(self.remedies_file, index=False, engine='openpyxl')
            logging.info(f"New medicine added: {med_name}")

    # ---------------- Update border sliders ----------------
    def update_border(self):
        self.top_offset = self.slider_top.value() / 100
        self.border_width = self.slider_width.value() / 100
        self.border_height = self.slider_height.value() / 100

    # ---------------- Generate PDF ----------------
    def print_label(self):
        try:
            med_name = self.medicine_search.text().upper()
            potency = self.potency_input.currentText().upper()
            dose = self.dose_input.text()
            quantity = self.quantity_input.text()
            time_val = self.time_input.text()
            shop = self.shop_input.text().upper()
            branch_phone = self.branch_phone_input.text().upper()

            if not med_name:
                QtWidgets.QMessageBox.warning(self, "Error", "Medicine name is required")
                return

            # Save autocomplete for potency, dose, quantity, time, shop
            for field, value in [("potency", potency), ("dose", dose),
                                 ("quantity", quantity), ("time", time_val),
                                 ("shop", shop)]:
                if value:
                    lst = self.autocomplete_data.setdefault(field, [])
                    if value not in lst:
                        lst.append(value)
            with open(autocomplete_file, "w") as f:
                json.dump(self.autocomplete_data, f)

            # Save record in single Excel
            record = {
                "Medicine": med_name,
                "Potency": potency,
                "Dose": dose,
                "Quantity": quantity,
                "Time": time_val,
                "Shop": shop,
                "Branch/Phone": branch_phone
            }
            if os.path.exists(records_file):
                df_records = pd.read_excel(records_file, engine="openpyxl")
                df_records = pd.concat([df_records, pd.DataFrame([record])], ignore_index=True)
            else:
                df_records = pd.DataFrame([record])
            df_records.to_excel(records_file, index=False, engine="openpyxl")

            # Generate PDF
            now = datetime.datetime.now()
            timestamp = now.strftime("%Y-%m-%d_%H-%M-%S")
            pdf_path = os.path.join(records_folder, f"label_45x25_{timestamp}.pdf")

            c = canvas.Canvas(pdf_path, pagesize=(1.77*inch, 0.98*inch))
            c.setLineWidth(1)
            c.rect(self.border_x*inch, self.border_y*inch, self.border_width*inch, self.border_height*inch)

            y = self.top_offset*inch
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString((self.border_x+self.border_width/2)*inch, y, f"{med_name} {potency}")

            y -= 0.15*inch
            c.setFont("Helvetica", 8)
            c.drawCentredString((self.border_x+self.border_width/2)*inch, y, f"{quantity}")

            y -= 0.15*inch
            c.drawCentredString((self.border_x+self.border_width/2)*inch, y, f"{dose}   {time_val}")

            y -= 0.15*inch
            c.setFont("Helvetica-Bold", 7)
            c.drawCentredString((self.border_x+self.border_width/2)*inch, y, f"{shop}")

            y -= 0.12*inch
            c.drawCentredString((self.border_x+self.border_width/2)*inch, y, f"{branch_phone}")

            c.save()

            # Open PDF
            if sys.platform == "win32":
                os.startfile(pdf_path)
            else:
                os.system(f"open {pdf_path}")

            self.status.setText("Label created, PDF saved, record appended.")

        except Exception as e:
            logging.error(f"Print/Save failed: {e}")
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed: {e}")
            self.status.setText(f"Error: {e}")


# ---------------- Run App ----------------
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = HomeoLabelApp()
    w.show()
    sys.exit(app.exec_())
