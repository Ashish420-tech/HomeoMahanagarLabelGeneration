import sys
import os
import json
import logging
import tempfile
import win32print
import win32api
from PyQt5 import QtWidgets, QtGui, QtCore
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from pdf2image import convert_from_path

# --- Logging setup ---
logging.basicConfig(filename="print_log.txt", level=logging.ERROR,
                    format="%(asctime)s - %(levelname)s - %(message)s")

# --- Constants ---
LABEL_WIDTH_MM = 50
LABEL_HEIGHT_MM = 30
PRINTER_NAME = "SNBC TVSE LP 46 NEO BPLE"


class LabelPrinterApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Homeopathy Label Printer")
        self.setGeometry(100, 100, 900, 600)
        self.pdf_path = None
        self.preview_mode = True  # 1st click = preview, 2nd = print
        self.font_size = 8  # default font size
        self.init_ui()

    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # --- Input fields ---
        form = QtWidgets.QFormLayout()
        self.name_input = QtWidgets.QLineEdit()
        self.time_input = QtWidgets.QLineEdit()
        self.potency_input = QtWidgets.QLineEdit()
        self.dose_input = QtWidgets.QLineEdit()
        self.branch_input = QtWidgets.QLineEdit()
        self.phone_input = QtWidgets.QLineEdit()

        form.addRow("Medicine Name:", self.name_input)
        form.addRow("Time:", self.time_input)
        form.addRow("Potency:", self.potency_input)
        form.addRow("Dose:", self.dose_input)
        form.addRow("Branch:", self.branch_input)
        form.addRow("Phone:", self.phone_input)

        layout.addLayout(form)

        # --- Settings: Top offset slider + Font size spin box ---
        settings_layout = QtWidgets.QHBoxLayout()

        settings_layout.addWidget(QtWidgets.QLabel("Top Offset (mm):"))
        self.offset_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.offset_slider.setRange(0, 20)
        self.offset_slider.setValue(5)
        settings_layout.addWidget(self.offset_slider)

        settings_layout.addWidget(QtWidgets.QLabel("Font Size:"))
        self.font_spin = QtWidgets.QSpinBox()
        self.font_spin.setRange(6, 20)
        self.font_spin.setValue(self.font_size)
        self.font_spin.valueChanged.connect(self.change_font_size)
        settings_layout.addWidget(self.font_spin)

        layout.addLayout(settings_layout)

        # --- Buttons ---
        button_layout = QtWidgets.QHBoxLayout()
        self.print_button = QtWidgets.QPushButton("Print Label")
        self.direct_button = QtWidgets.QPushButton("Direct Print")
        button_layout.addWidget(self.print_button)
        button_layout.addWidget(self.direct_button)
        layout.addLayout(button_layout)

        # --- Preview Area ---
        self.preview = QtWidgets.QGraphicsView()
        self.scene = QtWidgets.QGraphicsScene()
        self.preview.setScene(self.scene)
        layout.addWidget(self.preview)

        # --- Status ---
        self.status = QtWidgets.QLabel("")
        layout.addWidget(self.status)

        # --- Signals ---
        self.print_button.clicked.connect(self.handle_print_label)
        self.direct_button.clicked.connect(self.direct_print_label)

    def change_font_size(self, value):
        self.font_size = value

    def create_pdf(self):
        """Generate a new label PDF."""
        top_offset = self.offset_slider.value() * mm
        tmp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        c = canvas.Canvas(tmp_pdf.name, pagesize=(LABEL_WIDTH_MM * mm, LABEL_HEIGHT_MM * mm))
        c.translate(0, top_offset)

        x, y = 5 * mm, LABEL_HEIGHT_MM * mm - (10 * mm + top_offset)
        lines = [
            f"Name: {self.name_input.text()}",
            f"Time: {self.time_input.text()}",
            f"Potency: {self.potency_input.text()}",
            f"Dose: {self.dose_input.text()}",
            f"Branch: {self.branch_input.text()}",
            f"Phone: {self.phone_input.text()}",
        ]

        for line in lines:
            c.setFont("Helvetica", self.font_size)
            c.drawString(x, y, line)
            y -= 5 * mm

        c.showPage()
        c.save()
        self.pdf_path = tmp_pdf.name
        return self.pdf_path

    def show_pdf_preview(self, pdf_path):
        """Convert PDF to image and show in preview panel."""
        try:
            images = convert_from_path(pdf_path, dpi=150)
            if images:
                qt_image = QtGui.QImage(images[0].tobytes(), images[0].width, images[0].height,
                                        images[0].width * 3, QtGui.QImage.Format_RGB888)
                pixmap = QtGui.QPixmap.fromImage(qt_image)
                self.scene.clear()
                self.scene.addPixmap(pixmap)
                self.preview.fitInView(self.scene.sceneRect(), QtCore.Qt.KeepAspectRatio)
        except Exception as e:
            self.status.setText("⚠️ Could not render preview.")
            logging.error(f"Preview render failed: {e}")

    def handle_print_label(self):
        """Preview first, print second click."""
        pdf_file = self.create_pdf()
        self.show_pdf_preview(pdf_file)

        if self.preview_mode:
            self.status.setText("✅ Preview ready. Click again to print.")
            self.preview_mode = False
        else:
            self.preview_mode = True
            self.direct_print_label()

    def direct_print_label(self):
        """Direct print using system printer."""
        if not self.pdf_path:
            self.status.setText("⚠️ Please create a label first.")
            return

        try:
            win32api.ShellExecute(
                0,
                "printto",
                self.pdf_path,
                f'"{PRINTER_NAME}"',
                ".",
                0
            )
            self.status.setText(f"🖨️ Sent to printer: {PRINTER_NAME}")
        except Exception as e:
            self.status.setText("❌ Direct print failed.")
            logging.error(f"Direct print failed: {e}")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = LabelPrinterApp()
    window.show()
    sys.exit(app.exec_())
