import sys
import os
import json
import logging
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QCompleter, QTableWidgetItem, QSlider, QLabel, QHBoxLayout, QVBoxLayout
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# ------------------ Logging ------------------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ------------------ Homeopathy Label App ------------------
class HomeoLabelApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🏥 Homeopathy Label Generator")
        self.resize(900, 700)

        # ------------------ Paths ------------------
        self.records_folder = os.path.join(os.getcwd(), "records")
        os.makedirs(self.records_folder, exist_ok=True)
        self.excel_file = os.path.join(self.records_folder, 'records.xlsx')
        self.autocomplete_file = os.path.join(self.records_folder, 'autocomplete.json')

        # ------------------ Label settings ------------------
        self.label_width = 48 * mm / mm  # 48 mm
        self.label_height = 28 * mm / mm # 28 mm
        self.top_offset = 5 * mm
        self.font_size_med = 10
        self.label_generated = False

        # ------------------ Load data ------------------
        self.load_records()
        self.load_autocomplete()

        # ------------------ Build UI ------------------
        self.init_ui()

    # ------------------ Load existing records ------------------
    def load_records(self):
        if os.path.exists(self.excel_file):
            self.df_records = pd.read_excel(self.excel_file, engine='openpyxl')
        else:
            self.df_records = pd.DataFrame(columns=["Medicine","Potency","Dose","Quantity","Time","Shop","Branch/Phone"])

    # ------------------ Load autocomplete ------------------
    def load_autocomplete(self):
        if os.path.exists(self.autocomplete_file):
            try:
                self.autocomplete_data = json.load(open(self.autocomplete_file))
            except:
                self.autocomplete_data = {}
        else:
            self.autocomplete_data = {}
        for field in ["potency","dose","time","branch"]:
            if field not in self.autocomplete_data:
                self.autocomplete_data[field] = []

    # ------------------ Build UI ------------------
    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # Selected Medicine
        self.selected_medicine_label = QLabel("MEDICINE: ")
        self.selected_medicine_label.setStyleSheet("font-weight:bold; font-size:12pt;")
        layout.addWidget(self.selected_medicine_label)

        # Form Layout
        form = QtWidgets.QFormLayout()
        layout.addLayout(form)

        # Medicine
        self.medicine_input = QtWidgets.QLineEdit()
        self.medicine_input.setPlaceholderText("Enter medicine name")
        self.medicine_input.textChanged.connect(self.update_preview)
        form.addRow("Medicine:", self.medicine_input)

        # Potency
        self.potency_input = QtWidgets.QComboBox()
        self.potency_input.setEditable(True)
        self.potency_input.addItems(self.autocomplete_data.get("potency",[]))
        self.potency_input.setCompleter(QCompleter(self.autocomplete_data.get("potency",[])))
        self.potency_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Potency:", self.potency_input)

        # Dose
        self.dose_input = QtWidgets.QComboBox()
        self.dose_input.setEditable(True)
        self.dose_input.addItems(self.autocomplete_data.get("dose",[]))
        self.dose_input.setCompleter(QCompleter(self.autocomplete_data.get("dose",[])))
        self.dose_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Dose:", self.dose_input)

        # Quantity
        self.quantity_input = QtWidgets.QLineEdit()
        self.quantity_input.setPlaceholderText("Enter quantity")
        self.quantity_input.textChanged.connect(self.update_preview)
        form.addRow("Quantity:", self.quantity_input)

        # Time
        self.time_input = QtWidgets.QComboBox()
        self.time_input.setEditable(True)
        self.time_input.addItems(self.autocomplete_data.get("time",[]))
        self.time_input.setCompleter(QCompleter(self.autocomplete_data.get("time",[])))
        self.time_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Time:", self.time_input)

        # Shop
        self.shop_input = QtWidgets.QLineEdit()
        self.shop_input.setPlaceholderText("Enter Shop Name")
        self.shop_input.textChanged.connect(self.update_preview)
        form.addRow("Shop:", self.shop_input)

        # Branch/Phone
        self.branch_input = QtWidgets.QComboBox()
        self.branch_input.setEditable(True)
        self.branch_input.addItems(self.autocomplete_data.get("branch",[]))
        self.branch_input.setCompleter(QCompleter(self.autocomplete_data.get("branch",[])))
        self.branch_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Branch/Phone:", self.branch_input)

        # ------------------ Label Preview ------------------
        self.preview_frame = QtWidgets.QFrame()
        self.preview_frame.setFrameShape(QtWidgets.QFrame.Box)
        self.preview_frame.setFixedSize(400,180)
        preview_layout = QVBoxLayout(self.preview_frame)

        self.preview_line1 = QLabel("")
        self.preview_line1.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_line1.setStyleSheet("font-weight:bold; font-size:10pt;")
        preview_layout.addWidget(self.preview_line1)

        self.preview_line2 = QLabel("")
        self.preview_line2.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line2)

        self.preview_line3 = QLabel("")
        self.preview_line3.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line3)

        self.preview_line4 = QLabel("")
        self.preview_line4.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line4)

        self.preview_line5 = QLabel("")
        self.preview_line5.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line5)

        # ------------------ Top Offset Slider ------------------
        slider_layout = QHBoxLayout()
        slider_label = QLabel("Top Offset:")
        self.top_offset_slider = QSlider(QtCore.Qt.Horizontal)
        self.top_offset_slider.setMinimum(0)
        self.top_offset_slider.setMaximum(50)
        self.top_offset_slider.setValue(int(self.top_offset))
        self.top_offset_slider.valueChanged.connect(self.change_top_offset)
        slider_layout.addWidget(slider_label)
        slider_layout.addWidget(self.top_offset_slider)
        layout.addLayout(slider_layout)

        layout.addWidget(self.preview_frame)

        # ------------------ Print Button ------------------
        self.print_btn = QtWidgets.QPushButton("Print Label")
        self.print_btn.clicked.connect(self.print_label)
        layout.addWidget(self.print_btn)

        # Status
        self.status = QLabel("Ready")
        layout.addWidget(self.status)

    # ------------------ Change top offset ------------------
    def change_top_offset(self,value):
        self.top_offset = value
        self.update_preview()

    # ------------------ Update Preview ------------------
    def update_preview(self):
        med = self.medicine_input.text().upper()
        pot = self.potency_input.currentText().upper()
        dose = self.dose_input.currentText()
        qty = self.quantity_input.text()
        time_val = self.time_input.currentText()
        shop = self.shop_input.text().upper()
        branch = self.branch_input.currentText().upper()

        self.preview_line1.setText(f"{med} {pot}")
        self.preview_line2.setText(f"{qty}")
        self.preview_line3.setText(f"{dose}   {time_val}")
        self.preview_line4.setText(f"{shop}")
        self.preview_line5.setText(f"{branch}")

    # ------------------ Print / Generate PDF ------------------
    def print_label(self):
        try:
            med = self.medicine_input.text().upper()
            pot = self.potency_input.currentText().upper()
            dose = self.dose_input.currentText()
            qty = self.quantity_input.text()
            time_val = self.time_input.currentText()
            shop = self.shop_input.text().upper()
            branch = self.branch_input.currentText().upper()

            if not med:
                QtWidgets.QMessageBox.warning(self,"Error","Medicine name is required")
                return

            pdf_file = os.path.join(self.records_folder,"label.pdf")

            if not getattr(self,"label_generated",False):
                # ------------------ Generate PDF ------------------
                c = canvas.Canvas(pdf_file,pagesize=(48*mm,28*mm))
                c.setLineWidth(1)
                c.rect(2*mm,2*mm,48*mm-4*mm,28*mm-4*mm)

                y = 28*mm - self.top_offset
                c.setFont("Helvetica-Bold",10)
                c.drawCentredString(24*mm, y,f"{med} {pot}")

                y -= 5*mm
                c.setFont("Helvetica",8)
                c.drawCentredString(24*mm, y,f"{qty}")

                y -= 5*mm
                c.drawCentredString(24*mm, y,f"{dose}   {time_val}")

                y -= 5*mm
                c.setFont("Helvetica-Bold",7)
                c.drawCentredString(24*mm, y,f"{shop}")

                y -= 4*mm
                c.drawCentredString(24*mm, y,f"{branch}")

                c.save()
                self.label_generated = True

                # Open PDF preview
                if sys.platform=="win32":
                    os.startfile(pdf_file)
                else:
                    os.system(f"open {pdf_file}")
                self.status.setText("PDF generated. Click 'Print Label' again to print.")

                # ------------------ Save record & autocomplete ------------------
                record = {"Medicine":med,"Potency":pot,"Dose":dose,"Quantity":qty,"Time":time_val,"Shop":shop,"Branch/Phone":branch}
                self.df_records = pd.concat([self.df_records,pd.DataFrame([record])],ignore_index=True)
                self.df_records.to_excel(self.excel_file,index=False,engine='openpyxl')

                # Save autocomplete
                for field,value in [("potency",pot),("dose",dose),("time",time_val),("branch",branch)]:
                    if value and value not in self.autocomplete_data[field]:
                        self.autocomplete_data[field].append(value)
                with open(self.autocomplete_file,"w") as f:
                    json.dump(self.autocomplete_data,f)
            else:
                # ------------------ Send to Printer ------------------
                printer_name = "LP-46 Neo"
                try:
                    if sys.platform=="win32":
                        os.system(f'AcroRd32.exe /t "{pdf_file}" "{printer_name}"')
                    else:
                        os.system(f'lp -d "{printer_name}" "{pdf_file}"')
                    logging.info(f"Label sent to printer: {printer_name}")
                    self.status.setText(f"Label sent to printer: {printer_name}")
                except Exception as e:
                    logging.error(f"Printing failed: {e}")
                    QtWidgets.QMessageBox.critical(self,"Print Error",f"Printing failed: {e}")
                    self.status.setText(f"Printing failed: {e}")

        except Exception as e:
            logging.error(f"Print/Save failed: {e}")
            QtWidgets.QMessageBox.critical(self,"Error",f"Failed: {e}")
            self.status.setText(f"Error: {e}")

# ------------------ Run App ------------------
if __name__=="__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = HomeoLabelApp()
    w.show()
    sys.exit(app.exec_())
