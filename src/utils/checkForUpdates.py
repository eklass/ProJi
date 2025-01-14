import subprocess
import os

from utils.Constants import REPO_STATUS_CELL, VBA_SETTINGS_SHEET_NAME
from utils.excelLoader import ExcelLoader

global_excel_loader = None

def check_for_updates(project_root):
    setup_excel(VBA_SETTINGS_SHEET_NAME)
    os.chdir(project_root)  # Arbeitsverzeichnis ändern

    # Prüfen, ob es Updates gibt
    subprocess.run(["git", "fetch"], check=True)
    result = subprocess.run(["git", "status", "-uno"], capture_output=True, text=True)
    if "behind" in result.stdout:
        get_excel_loader().log_to_excel("Ein Update ist verfügbar!\nBitte aktualisieren Sie Ihr Repository.", True)
        get_excel_loader().vba_settings_sheet.range(REPO_STATUS_CELL).value = 'behind'
        show_applescript_popup()
    else:
        get_excel_loader().log_to_excel("ProJi Repo ist up to date.", True)
        get_excel_loader().vba_settings_sheet.range(REPO_STATUS_CELL).value = 'latest'


def show_applescript_popup():
    script = """
    display dialog "Ein Update ist verfügbar!\nBitte aktualisieren Sie Ihr Repository." buttons {"OK"} default button "OK"
    """
    subprocess.run(["osascript", "-e", script], check=True)


def setup_excel(sheet_name):
    excel_loader = ExcelLoader()
    excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)


def set_excel_loader(excel_loader):
    global global_excel_loader
    if not isinstance(excel_loader, ExcelLoader):
        raise TypeError("Das übergebene Objekt ist kein ExcelLoader.")
    global_excel_loader = excel_loader


def get_excel_loader():
    global global_excel_loader
    if global_excel_loader is None:
        raise ValueError("ExcelLoader ist noch nicht gesetzt. Bitte setzen Sie ihn zuerst.")
    return global_excel_loader



if __name__ == "__main__":
    # Aktuelles Verzeichnis speichern
    current_dir = os.getcwd()
    # Zwei Ebenen hochgehen
    project_root = os.path.abspath(os.path.join(current_dir, "../../"))
    check_for_updates(project_root)
