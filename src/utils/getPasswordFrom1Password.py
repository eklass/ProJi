import datetime
import os
import shutil
import subprocess
import sys
import time

from utils.Constants import VBA_SHEET_PASSWORD_TOKEN_TIMESTAMP_CELL, VBA_SHEET_STATUS_OF_SESSION_TOKEN_CELL, LOG_ONE_PASSWORD_ROUTINE_CELL, WEEKDAY_CONSOLE_OUTPUT_CELL, \
    VBA_SHEET_SESSION_TOKEN_CELL
from utils.excelLoader import ExcelLoader

# Pfad zur Datei, wo das Session-Token gespeichert werden soll
log_one_password_routine = None

output_messages = []
global_excel_loader = None

def find_op_path():
    op_path = shutil.which('op')
    if not op_path:
        # Fallback auf bekannten Standardpfad (falls vorhanden)
        possible_paths = [
            '/opt/homebrew/bin/op',  # Homebrew (macOS ARM64)
            '/usr/local/bin/op',    # Homebrew (Intel macOS)
            '/usr/bin/op'           # Standardpfade auf Linux
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
    return op_path


def add_1password_account(address, email, secret_key, master_password):
    # Dynamisch nach "op" suchen
    op_path = find_op_path()
    if not op_path:
        raise FileNotFoundError("Die OnePassword CLI 'op' wurde nicht gefunden. Bitte installieren Sie sie.")

    add_account_command = [
        op_path, 'account', 'add',
        '--address', address,
        '--email', email,
        '--secret-key', secret_key,
        '--raw'
    ]

    try:
        # Führt den Befehl zum Hinzufügen des Kontos aus, mit Master-Passwort als Eingabe
        process = subprocess.Popen(
            add_account_command,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        # Übergibt das Passwort und liest die Ausgabe des Subprozesses
        output, error = process.communicate(input=rf'{master_password}' + '\n')

        if process.returncode != 0:
            get_excel_loader().log_to_excel("Fehler beim Hinzufügen des Accounts: " + error.strip())
            sys.exit(1)

    except Exception as e:
        get_excel_loader().log_to_excel("Fehler beim Ausführen von subprocess.Popen in add_1password_account: " + str(e))
        sys.exit(1)

    get_excel_loader().log_to_excel("Account erfolgreich hinzugefügt.")


def save_session_token(token):
    get_excel_loader().vba_settings_sheet.range(VBA_SHEET_SESSION_TOKEN_CELL).value = token
    current_time = datetime.datetime.now().strftime("%H:%M:%S")  # Formatierung der Uhrzeit
    # Eintragen der aktuellen Uhrzeit in die Zelle K5
    get_excel_loader().vba_settings_sheet.range(VBA_SHEET_PASSWORD_TOKEN_TIMESTAMP_CELL).value = current_time
    get_excel_loader().vba_settings_sheet.range(VBA_SHEET_STATUS_OF_SESSION_TOKEN_CELL).value = 'valid'
    get_excel_loader().wb.save()


def load_session_token():
    return get_excel_loader().vba_settings_sheet.range(VBA_SHEET_SESSION_TOKEN_CELL).value


def remove_session_token():
    get_excel_loader().vba_settings_sheet.range(VBA_SHEET_SESSION_TOKEN_CELL).value = ''


def is_session_valid(sheet_name, session_token):
    global log_one_password_routine
    setup_excel(sheet_name)
    if not session_token:
        get_excel_loader().log_to_excel("No valid sessionToken")
        get_excel_loader().vba_settings_sheet.range(VBA_SHEET_STATUS_OF_SESSION_TOKEN_CELL).value = 'invalid'
        get_excel_loader().wb.save()
        return False

    log_one_password_routine = get_excel_loader().vba_settings_sheet.range(LOG_ONE_PASSWORD_ROUTINE_CELL).value

    # Verwendet den `op whoami`-Befehl, um die Gültigkeit der Session zu überprüfen
    op_path = find_op_path()
    if not op_path:
        raise FileNotFoundError("Die OnePassword CLI 'op' wurde nicht gefunden. Bitte installieren Sie sie.")

    test_command = [op_path, 'whoami', '--session', session_token]

    result = subprocess.run(test_command, capture_output=True, text=True)
    session_valid = result.returncode == 0
    if not session_valid:
        remove_session_token()
        get_excel_loader().log_to_excel("Cleaned Up Session Token")
        get_excel_loader().vba_settings_sheet.range(VBA_SHEET_STATUS_OF_SESSION_TOKEN_CELL).value = 'invalid'
    else:
        get_excel_loader().vba_settings_sheet.range(VBA_SHEET_STATUS_OF_SESSION_TOKEN_CELL).value = 'valid'
    get_excel_loader().wb.save()
    get_excel_loader().wb.activate(True)
    return session_valid


def sign_in_to_1password(sheet_name, master_password):

    session_token = load_session_token()
    if session_token and is_session_valid(sheet_name, session_token):
        masked_token = session_token[:5] + '*****' + session_token[-5:]
        get_excel_loader().log_to_excel(f"Session ist gültig. Verwende existierendes Session-Token: {masked_token}")
        return session_token
    else:
        get_excel_loader().log_to_excel("Session ist nicht gültig. Hole neuen Token")

    op_path = find_op_path()
    if not op_path:
        raise FileNotFoundError("Die OnePassword CLI 'op' wurde nicht gefunden. Bitte installieren Sie sie.")

    sign_in_command = [op_path, 'signin', '--raw']

    get_excel_loader().log_to_excel("op sign in is called ...")

    try:
        # Erstellen eines Subprozesses mit Popen und Bereitstellung der Passworteingabe
        process = subprocess.Popen(
            sign_in_command,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )

        # Übergeben des Passworts an den signin-Befehl
        new_session_token, error = process.communicate(input=master_password + '\n')
        masked_token = new_session_token[:5] + '*****' + new_session_token[-5:]
        get_excel_loader().log_to_excel(f"Neuer Session-Token: {masked_token}")

        if process.returncode != 0:
            get_excel_loader().log_to_excel("Fehler beim Anmelden: " + error.strip())
            sys.exit(1)

    except Exception as e:
        get_excel_loader().log_to_excel("Fehler beim Ausführen von subprocess.Popen: " + str(e))
        sys.exit(1)

    save_session_token(new_session_token)
    return new_session_token


# Funktion zum Abrufen eines Items
def get_credentials(sheet_name, password_reference):
    setup_excel(sheet_name)

    session_token = read_session_token()
    if password_reference is None or session_token is None:
        get_excel_loader().log_to_excel("Session-Token oder Password-Reference ist nicht gesetzt.")
        sys.exit(1)

    op_path = find_op_path()
    if not op_path:
        raise FileNotFoundError("Die OnePassword CLI 'op' wurde nicht gefunden. Bitte installieren Sie sie.")

    get_item_command = [op_path, 'item', 'get', password_reference, '--session', session_token]
    try:
        item_result = subprocess.run(
            get_item_command,
            capture_output=True,
            text=True
        )
        item_result.check_returncode()  # Dies prüft, ob der Befehl erfolgreich war und löst eine Ausnahme aus, wenn nicht.
    except subprocess.CalledProcessError as e:
        get_excel_loader().log_to_excel("Fehler beim Abrufen des Items: " + e.stderr.strip())
        sys.exit(1)

    # Parsen der Ausgabe, um die ID des Items zu extrahieren
    item_id = None
    lines = item_result.stdout.split('\n')
    for line in lines:
        if 'ID:' in line:
            item_id = line.split()[1].strip()
            break

    if not item_id:
        get_excel_loader().log_to_excel("Item ID konnte nicht gefunden werden.")
        sys.exit(1)

    op_path = find_op_path()
    if not op_path:
        raise FileNotFoundError("Die OnePassword CLI 'op' wurde nicht gefunden. Bitte installieren Sie sie.")

    # Zweiter Befehl, um die Credential-Details zu enthüllen
    get_password_command = [op_path, 'item', 'get', item_id, '--session', session_token, '--fields', 'username,password', '--reveal']
    credentials_result = subprocess.run(
        get_password_command,
        capture_output=True,
        text=True
    )

    try:
        credentials_result.check_returncode()
    except subprocess.CalledProcessError as e:
        get_excel_loader().log_to_excel("Fehler beim Abrufen der Credentials: " + e.stderr.strip())
        sys.exit(1)

    # Parsen der Ausgabe, um Benutzername und Passwort zu extrahieren
    credentials = credentials_result.stdout.strip().split(',', 1)
    if len(credentials) < 1:
        get_excel_loader().log_to_excel("Benutzername und Passwort konnten nicht korrekt extrahiert werden.")
        sys.exit(1)

    username, password = credentials

    if(',' in password or '"' in password):
        password = password[1:-1]
        password = password.replace('""', '"')

    if not password:
        get_excel_loader().log_to_excel("Passwort konnte nicht gefunden werden.")
        sys.exit(1)

    return {"username": username, "password": password}


def get_or_create_session_token(sheet_name, master_password, secret_key, email, address):
    setup_excel(sheet_name)

    get_excel_loader().log_to_excel("Starting Get Password Routine for 1Password ...")
    get_excel_loader().log_to_excel(f"Account Address from VBA-Settings: {address}")
    get_excel_loader().log_to_excel(f"Email from VBA-Settings: {email}")
    if address and email and secret_key and master_password:
        # Account hinzufügen
        add_1password_account(address, email, secret_key, master_password)
    else:
        # Argumente fehlen, gib eine Fehlermeldung aus
        get_excel_loader().log_to_excel("Fehler: Alle Parameter müssen gefüllt sein.")
        sys.exit(1)
    # Anmelden und Session Token erhalten
    session_token = sign_in_to_1password(sheet_name, master_password)
    return session_token


def clear_log(sheet):
    sheet.range(WEEKDAY_CONSOLE_OUTPUT_CELL).value = ''


def read_session_token():
    # Versuche für bis zu 10 Sekunden den session_token zu lesen
    timeout = time.time() + 10  # 10 Sekunden ab jetzt
    session_token = None

    while time.time() < timeout:
        # Lese den Wert der Zelle
        session_token = get_excel_loader().vba_settings_sheet.range(VBA_SHEET_SESSION_TOKEN_CELL).value
        if session_token:
            break
        time.sleep(0.5)  # Warte 0.5 Sekunden, bevor die nächste Überprüfung stattfindet

    return session_token

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


if __name__ == '__main__':

    excel_loader = ExcelLoader()
    excel_loader.load_excel('Dienstag')
    set_excel_loader(excel_loader)

    get_excel_loader().wb.app.display_alerts = False
    log_one_password_routine = get_excel_loader().vba_settings_sheet.range(LOG_ONE_PASSWORD_ROUTINE_CELL).value
    get_or_create_session_token('Dienstag', r'MASTER_PASSWORD', r'SECRET_KEY', 'EMAIL', 'ONE_PASSWORD_URL')

