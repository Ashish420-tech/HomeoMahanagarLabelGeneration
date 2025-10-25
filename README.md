# 🏥 Homeopathy Label Generator

A professional desktop application for generating and printing medicine labels for homeopathy pharmacies. Built with Python and PyQt5, this tool streamlines the label creation process with an intuitive interface, autocomplete features, and direct printing capabilities.

![Python](https://img.shields.io/badge/Python-3.12-blue.svg)
![PyQt5](https://img.shields.io/badge/PyQt5-GUI-green.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

---

## ✨ Features

### 🔍 **Smart Medicine Search**
- Real-time search for medicines by Latin or common names
- Auto-suggestion table with instant results
- Add new medicines on-the-fly

### 📝 **Label Customization**
- Adjustable font sizes for different label requirements
- Live preview of labels before printing
- Support for custom potencies, doses, and timing
- Shop name and branch/phone customization

### 🖨️ **Flexible Printing**
- Direct print to any connected USB/network printer
- PDF preview generation
- Automatic printer detection and status checking
- 50mm x 30mm label size (configurable)

### 💾 **Data Management**
- Auto-save all printed labels to Excel
- Autocomplete for frequently used values (potencies, doses, times)
- Persistent storage of shop names and branches
- Error logging for troubleshooting

### 🎨 **User Interface**
- Responsive design with DPI awareness
- Full-screen maximized view
- Clean, professional layout
- Real-time label preview

---

## 📋 Requirements

- **Python**: 3.8 or higher
- **Operating System**: Windows 10/11
- **Dependencies**:
  - PyQt5
  - pandas
  - openpyxl
  - reportlab
  - pywin32

---

## 🚀 Installation

### Option 1: Run from Source

1. **Clone the repository:**
git clone https://github.com/Ashish420-tech/HomeoMahanagarLabelGeneration.git
cd HomeoMahanagarLabelGeneration

2. **Install dependencies:**
pip install PyQt5 pandas openpyxl reportlab pywin32


3. **Run the application:**
python HomeoLabelApp.py


### Option 2: Use Pre-built Executable

1. Download the latest `.exe` from the [Releases](https://github.com/Ashish420-tech/HomeoMahanagarLabelGeneration/releases) page
2. Extract to a folder
3. Ensure `remedies.xlsx` is in the same directory
4. Double-click `HomeoLabelApp.exe`

---

## 🛠️ Building from Source

To create your own executable:

1. **Install PyInstaller:**

2. **Build the executable:**
pyinstaller --clean --noconfirm --onefile --windowed --name "HomeopathyLabelGenerator" --add-data "remedies.xlsx;." HomeoLabelApp.py

3. **Find the executable:**
dist/HomeopathyLabelGenerator.exe


---

## 📖 Usage Guide

### 1. **Search for a Medicine**
- Type the medicine name (Latin or common) in the search box
- Click on a suggestion from the table to select it
- Or click "Add New Medicine" to add unlisted medicines

### 2. **Configure Label Details**
- **Potency**: Select or type (e.g., 30C, 200, 1M)
- **Dose**: Enter dosage instructions (e.g., "4 drops", "2 pills")
- **Time**: Specify timing (e.g., "3 times daily", "before meals")
- **Shop Name**: Your pharmacy name
- **Branch/Phone**: Branch location or contact number

### 3. **Preview and Print**
- See live preview in the preview panel
- Adjust font size if needed
- Click **"Preview PDF"** to view before printing
- Click **"Direct Print"** to send to printer immediately

### 4. **Records**
- All labels are automatically saved to `records/records.xlsx`
- PDFs are stored in the `records/` folder
- Error logs available in `records/error_log.txt`

---
HomeoMahanagarLabelGeneration/
├── HomeoLabelApp.py # Main application file
├── remedies.xlsx # Medicine database (Latin + Common names)
├── records/ # Auto-generated folder
│ ├── records.xlsx # Saved label history
│ ├── autocomplete.json # Autocomplete data
│ ├── label.pdf # Latest generated label
│ └── error_log.txt # Error logs
├── build_exe.bat # Build script for creating .exe
├── .gitignore # Git ignore rules
└── README.md # This file
## 📁 Project Structure

---

## 🔧 Configuration

### Adding Medicines
Edit `remedies.xlsx` or use the "Add New Medicine" button in the app:
- **Column 1 (common_col)**: Common name (e.g., "Arnica")
- **Column 2 (latin_col)**: Latin name (e.g., "Arnica montana")

### Label Dimensions
Default: **50mm × 30mm**

To change, edit in `HomeoLabelApp.py`:
width_mm, height_mm = 50, 30 # Line ~XXX


### Printer Settings
- Use the "Refresh" button to detect new printers
- Printer status checked before printing
- Falls back to PDF preview if printing fails

---

## 🐛 Troubleshooting

### **Printer Not Detected**
- Ensure printer is ON and connected via USB
- Try a different USB port
- Print a Windows test page first
- Click "Refresh Printers" in the app

### **Module Not Found Error**
pip install --upgrade PyQt5 pandas openpyxl reportlab pywin32


### **Permission Denied on records.xlsx**
- Close Excel if `records.xlsx` is open
- The app buffers records and saves when possible

### **Application Not Opening**
- Check `records/error_log.txt` for details
- Run from command line to see error messages:
python HomeoLabelApp.py


---

## 🤝 Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

---

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 👨‍💻 Author

**Ashish Mondal**
- GitHub: [@Ashish420-tech](https://github.com/Ashish420-tech)
- Project: [HomeoMahanagarLabelGeneration](https://github.com/Ashish420-tech/HomeoMahanagarLabelGeneration)

---

## 🙏 Acknowledgments

- Built with [PyQt5](https://www.riverbankcomputing.com/software/pyqt/)
- PDF generation powered by [ReportLab](https://www.reportlab.com/)
- Data handling with [pandas](https://pandas.pydata.org/)
- Inspired by the needs of homeopathy practitioners

---

## 📞 Support

For issues, questions, or feature requests:
- Open an [Issue](https://github.com/Ashish420-tech/HomeoMahanagarLabelGeneration/issues)
- Email: mimokool2011@gmail.com

---

## 🔄 Version History

### v1.0.0 (2025-10-25)
- Initial release
- Medicine search with autocomplete
- Label preview and printing
- Excel record keeping
- DPI-aware responsive UI

---

**Made with ❤️ for Homeopathy Practitioners**


