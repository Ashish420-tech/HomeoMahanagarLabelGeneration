import subprocess
import sys

def install_reportlab():
    try:
        import reportlab
        print("✅ ReportLab is already installed.")
    except ImportError:
        print("⚠ ReportLab not found. Installing now...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "reportlab"])
        try:
            import reportlab
            print("✅ ReportLab installed successfully!")
        except ImportError:
            print("❌ Failed to install ReportLab. Please install manually.")

if __name__ == "__main__":
    install_reportlab()
