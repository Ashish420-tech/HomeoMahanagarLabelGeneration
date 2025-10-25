import sys
import os
import pandas as pd
from datetime import datetime
from rapidfuzz import process
from PyQt5 import QtWidgets
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch


class HomeoApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Homeopathy Remedy Label Printing")
        self.resize(750, 700)

        # Files
        self.remedies_file = "remedies.xlsx"
        self.records_file = "records.xlsx"

        # Create default remedies file if missing
        if not os.path.exists(self.remedies_file):
            df = pd.DataFrame({"Remedy": ["Arnica", "Belladonna", "Nux Vomica"]})
            df.to_excel(self.remedies_file, index=False)

        # Layout
        layout = QtWidgets.QVBoxLayout(self)

        # Search bar
        self.search_bar = QtWidgets.QLineEdit()
        self.search_bar.setPlaceholderText("Search remedies...")
        layout.addWidget(self.search_bar)

        # Search results
        self.table = QtWidgets.QTableWidget()
        self.table.setColumnCount(1)
        self.table.setHorizontalHeaderLabels(["Remedy"])
        layout.addWidget(self.table)

        # Extra Inputs
        form_layout = QtWidgets.QFormLayout()
        self.potency_input = QtWidgets.QLineEdit()
        self.dose_input = QtWidgets.QComboBox()
        self.dose_input.addItems(["1", "2", "3", "4"])
        self.shop_input = QtWidgets.QLineEdit()
        self.shop_input.setText("HOMEO MAHANAGR")
        self.quantity_input = QtWidgets.QLineEdit()
        self.quantity_input.setPlaceholderText("e.g., 2 drops or 3 globules")

        form_layout.addRow("Potency:", self.potency_input)
        form_layout.addRow("Dose (times/day):", self.dose_input)
        form_layout.addRow("Quantity (drops/globules):", self.quantity_input)
        form_layout.addRow("Shop Name:", self.shop_input)
        layout.addLayout(form_layout)

        # Time Inputs container
        self.time_container = QtWidgets.QVBoxLayout()
        layout.addWidget(QtWidgets.QLabel("Select Times:"))
        layout.addLayout(self.time_container)

        # Buttons
        btn_layout = QtWidgets.QHBoxLayout()
        self.search_btn = QtWidgets.QPushButton("Search")
        self.print_btn = QtWidgets.QPushButton("Print Label")
        btn_layout.addWidget(self.search_btn)
        btn_layout.addWidget(self.print_btn)
        layout.addLayout(btn_layout)

        # Actions
        self.search_btn.clicked.connect(self.search_remedies)
        self.print_btn.clicked.connect(self.print_label)
        self.dose_input.currentIndexChanged.connect(self.update_time_inputs)

        # Time input storage
        self.time_inputs = []

        # Initialize time inputs based on default dose
        self.update_time_inputs()

    # ---------------------- Time Inputs ----------------------
    def update_time_inputs(self):
        """Adjust the number of time inputs according to selected dose"""
        # Clear existing widgets
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

    # ---------------------- Search ----------------------
    def search_remedies(self):
        query = self.search_bar.text().strip()
        if not query:
            return
        df = pd.read_excel(self.remedies_file)
        remedies = df["Remedy"].tolist()
        results = process.extract(query, remedies, limit=10)

        self.table.setRowCount(len(results))
        for row, (match, score, idx) in enumerate(results):
            self.table.setItem(row, 0, QtWidgets.QTableWidgetItem(match))

    # ---------------------- Print & Save ----------------------
    def print_label(self):
        row = self.table.currentRow()
        if row < 0:
            QtWidgets.QMessageBox.warning(self, "No Remedy", "Please select a remedy from the table.")
            return

        medicine = self.table.item(row, 0).text()
        potency = self.potency_input.text().strip()
        dose = self.dose_input.currentText()
        quantity = self.quantity_input.text().strip()
        shop = self.shop_input.text().strip()

        # Collect times dynamically
        times = []
        for hour, minute, ampm in self.time_inputs:
            times.append(f"{hour.value():02d}:{minute.value():02d} {ampm.currentText()}")

        # ---------------------- Create Label PDF ----------------------
        label_file = "label.pdf"
        c = canvas.Canvas(label_file, pagesize=(2*inch, 1*inch))

        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(1*inch, 0.85*inch, medicine)

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

        # ---------------------- Print to Default Printer ----------------------
        try:
            if sys.platform == "win32":
                os.startfile(label_file, "print")
            else:
                QtWidgets.QMessageBox.information(self, "Label Saved", f"Label saved as {label_file}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", str(e))

        # ---------------------- Save Record ----------------------
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        record = {
            "Timestamp": timestamp,
            "Medicine": medicine,
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

        QtWidgets.QMessageBox.information(self, "Label Printed", "Label printed and record saved successfully!")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = HomeoApp()
    window.show()
    sys.exit(app.exec_())
