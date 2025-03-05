import subprocess
import sys


def install():
    subprocess.check_call([sys.executable, "-m", "pip", "install", "tk==0.1.0"])
    subprocess.check_call([sys.executable, "-m", "pip", "install", "smartsheet-python-sdk==3.0.4"])
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyautogui==0.9.53"])
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx==1.1.2"])


install()