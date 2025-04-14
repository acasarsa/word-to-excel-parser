# automation-project/automation_app/setup_script.py
import subprocess
import sys
import os

def install():
    print("📦 Installing requirements from requirements.txt...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])

    print("📁 Creating necessary folders...")
    os.makedirs("word_file", exist_ok=True)
    os.makedirs("excel_file", exist_ok=True)

    print("✅ Setup complete! You can now run `python run.py` or continue configuring.")
