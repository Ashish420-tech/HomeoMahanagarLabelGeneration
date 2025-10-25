`import sys
import os
import pandas as pd
from datetime import datetime
from PyQt5 import QtWidgets, QtCore
from rapidfuzz import process
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch
import win32api
import logging

# ----------------- Configure Logging -----------------
logging.basicConfig(
    filename="error_log.txt",
    filemode="a",  # append mode
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# ----------------- Main App -----------------
class HomeoLabelApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🏥 Homeopathy Label Printing")
        self.resize(800, 700)

        try:
            # Files
            self.remedies_file = "remedies.xlsx"
            self.records_file = "records.xlsx"

            # Load or create remedies Excel
            if not os.path.exists(self.remedies_file):
                df = pd.DataFrame({
                    "common_col": ["Arnica", "Belladonna", "Nux Vomica"],
                    "latin_col": ["Arnica montana", "Atropa belladonna", "Strychnos nux-vomica"]
                })
                df.to_excel(self.remedies_file, index=False)

            self.df_remedies = pd.read_excel(self.remedies_file)
            self.df_remedies.columns = [c.strip() for c in self.df_remedies.columns]  # clean spaces

            # ------------------- Layout -------------------
            layout = QtWidgets.QVBoxLayout(self)

            # Search bar
            self.search_bar = QtWidgets.QLineEdit()
            self.search_bar.setPlaceholderText("Search medicine by Common or Latin Name...")
            layout.addWidget(self.search_bar)

            # Search results table
            self.table = QtWidgets.QTableWidget()
            self.table.setColumnCount(2)
            self.table.setHorizontalHeaderLabels(["Common Name", "Latin Name"])
            self.table.horizontalHeader().setStretchLastSection(True)
            layout.addWidget(self.table)

            # Input fields
            form_layout = QtWidgets.QFormLayout()
            self.potency_input = QtWidgets.QLineEdit()
            self.dose_input = QtWidgets.QComboBox()
            self.dose_input.addItems(["1", "2", "3", "4"])
            self.quantity_input = QtWidgets.QLineEdit()
            self.quantity_input.setPlaceholderText("e.g., 2 drops or 3 globules")
            self.shop_input = QtWidgets.QLineEdit()
            self.shop_input.setText("HOMEO MAHANAGR")

            form_layout.addRow("Potency:", self.potency_input)
            form_layout.addRow("Dose (times/day):", self.dose_input)
            form_layout.addRow("Quantity (drops/globules):", self.quantity_input)
            form_layout.addRow("Shop Name:", self.shop_input)
            layout.addLayout(form_layout)

            # Time inputs container
            layout.addWidget(QtWidgets.QLabel("Times (12-hour format with AM/PM):"))
            self.time_container = QtWidgets.QVBoxLayout()
            layout.addLayout(self.time_container)
            self.time_inputs = []
            self.update_time_inputs()  # initialize

            # Buttons
            btn_layout = QtWidgets.QHBoxLayout()
            self.search_btn = QtWidgets.QPushButton("Search")
            self.print_btn = QtWidgets.QPushButton("Print Label")
            btn_layout.addWidget(self.search_btn)
            btn_layout.addWidget(self.print_btn)
            layout.addLayout(btn_layout)

            # ------------------- Connections -------------------
            self.search_btn.clicked.connect(self.search_medicine)
            self.dose_input.currentIndexChanged.connect(self.update_time_inputs)
            self.print_btn.clicked.connect(self.print_label)
        except Exception as e:
            logging.error("Error initializing HomeoLabelApp: %s", str(e))
            QtWidgets.QMessageBox.critical(self, "Error", f"Initialization failed: {e}")

    # ------------------- Dynamic Time Inputs -------------------
    def update_time_inputs(self):
        try:
            # Clear previous inputs
            for i in reversed(range(self.time_container.count())):
                widget = self.time_container.itemAt(i).widget()
                if widget:
                    widget.setParent(None)
            self.time_inputs = []

            dose = int(self.dose_input.currentText())
            for i in range(dose):
                row_layout = QtWidgets.QHBoxLayout()
                hour = QtWidgets.QSpinBox()
                hour.setRange(1, 12)
                hour.setValue(8)

                minute = QtWidgets.QSpinBox()
                minute.setRange(0, 59)
                minute.setValue(0)

                ampm = QtWidgets.QComboBox()
                ampm.addItems(["AM", "PM"])

                row_layout.addWidget(QtWidgets.QLabel(f"Time {i+1}:"))
                row_layout.addWidget(hour)
                row_layout.addWidget(QtWidgets.QLabel(":"))
                row_layout.addWidget(minute)
                row_layout.addWidget(ampm)

                container = QtWidgets.QWidget()
                container.setLayout(row_layout)
                self.time_container.addWidget(container)

                self.time_inputs.append((hour, minute, ampm))
        except Exception as e:
            logging.error("Error in update_time_inputs: %s", str(e))
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to update time inputs: {e}")

    # ------------------- Search Medicine -------------------
    def search_medicine(self):
        try:
            query = self.search_bar.text().strip().lower()
            if not query:
                return

            # Fuzzy search
            names = list(self.df_remedies["common_col"]) + list(self.df_remedies["latin_col"])
            matches = process.extract(query, names, limit=20)
            found = set([m[0] for m in matches if m[1] > 50])

            df_filtered = self.df_remedies[self.df_remedies.apply(
                lambda r: r["common_col"] in found or r["latin_col"] in found, axis=1
            )]

            self.table.setRowCount(len(df_filtered))
            for i, (_, row) in enumerate(df_filtered.iterrows()):
                self.table.setItem(i, 0, QtWidgets.QTableWidgetItem(str(row["common_col"])))
                self.table.setItem(i, 1, QtWidgets.QTableWidgetItem(str(row["latin_col"])))
        except Exception as e:
            logging.error("Error in search_medicine: %s", str(e))
            QtWidgets.QMessageBox.critical(self, "Error", f"Search failed: {e}")

    # ------------------- Print Label & Save Record -------------------
    def print_label(self):
        try:
            row = self.table.currentRow()
            if row < 0:
                QtWidgets.QMessageBox.warning(self, "No Medicine", "Please select a medicine from the table.")
                return

            # Get data
            common_name = self.table.item(row, 0).text()
            potency = self.potency_input.text().strip()
            dose = self.dose_input.currentText()
            quantity = self.quantity_input.text().strip()
            shop = self.shop_input.text().strip()

            times = []
            for hour, minute, ampm in self.time_inputs:
                times.append(f"{hour.value():02d}:{minute.value():02d} {ampm.currentText()}")

            # Create Label PDF
            label_file = "label.pdf"
            c = canvas.Canvas(label_file, pagesize=(2*inch, 1*inch))
            c.setFont("Helvetica-Bold", 9)
            c.drawCentredString(1*inch, 0.85*inch, common_name)

            c.setFont("Helvetica", 7)
            c.drawString(0.1*inch, 0.68*inch, f"Potency: {potency}")
            c.drawString(0.1*inch, 0.55*inch, f"Dose: {dose} times/day")
            c.drawString(0.1*inch, 0.42*inch, f"Quantity: {quantity}")
            c.drawString(0.1*inch, 0.29*inch, f"Shop: {shop}")

            y = 0.16 * inch
            for t in times:
                c.drawString(0.1*inch, y, f"Time: {t}")
                y -= 0.12 * inch
                if y < 0.05 * inch:
                    break

            c.showPage()
            c.save()

            # Open PDF in default viewer instead of direct printing
            if os.path.exists(label_file):
                win32api.ShellExecute(0, "open", label_file, None, ".", 1)

            # Save Record
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            record = {
                "Timestamp": timestamp,
                "Medicine": common_name,
                "Potency": potency,
                "Dose": dose,
                "Quantity": quantity,
                "Times": ",".join(times),
                "Shop": shop
            }

            if os.path.exists(self.records_file):
                df_records = pd.read_excel(self.records_file)
                df_records = pd.concat([df_records, pd.DataFrame([record])], ignore_index=True)
            else:
                df_records = pd.DataFrame([record])

            df_records.to_excel(self.records_file, index=False)

            QtWidgets.QMessageBox.information(self, "Success", "Label created! Opened in PDF viewer. Record saved.")
        except Exception as e:
            logging.error("Error in print_label: %s", str(e))
            QtWidgets.QMessageBox.critical(self, "Error", f"Print/Save failed: {e}")

# ----------------- Run Application -----------------
if __name__ == "__main__":
    try:
        app = QtWidgets.QApplication(sys.argv)
        window = HomeoLabelApp()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        logging.critical("Application crashed: %s", str(e))
`