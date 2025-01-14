import os
import queue
import shutil
import subprocess

from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select

from utils.getPasswordFrom1Password import get_credentials

from utils.Constants import PROJETRON_LOCALE_SETTING_CELL, PROJEKTRON_STATUS_COLUMN, PROJEKTRON_DOMAIN_CELL, \
    ONE_PASSWORD_REFERENCE_PROJEKTRON_CELL, HEADLESS_MODE_CELL
from utils.excelLoader import ExcelLoader

global_excel_loader = None


def wait_for_optional_element_to_be_clickable(driver, locator, duration=2):
    try:
        return WebDriverWait(driver, duration).until(EC.element_to_be_clickable(locator))
    except TimeoutException:
        return None

def wait_for_element_to_be_clickable(driver, locator, duration=10):
    return WebDriverWait(driver, duration).until(EC.element_to_be_clickable(locator))


def login_to_website(driver, email, password):
    full_projektron_url = open_time_booking_page_in_projektron(driver)
    get_excel_loader().log_to_excel(f'Log in to {full_projektron_url} ...')
    wait_for_element_to_be_clickable(driver, (By.CLASS_NAME, 'oAuthLoginLink')).click()

    # Warte, bis das E-Mail-Eingabefeld interaktionsbereit ist und gib die E-Mail ein
    email_input = wait_for_element_to_be_clickable(driver, (By.NAME, 'loginfmt'))
    email_input.send_keys(email)

    # Warte, bis der Weiter-Button klickbar ist, und klicke darauf
    next_button = wait_for_element_to_be_clickable(driver, (By.ID, 'idSIButton9'))
    next_button.click()

    # Warte, bis das Passwort-Eingabefeld interaktionsbereit ist und gib das Passwort ein
    password_input = wait_for_element_to_be_clickable(driver, (By.NAME, 'passwd'))
    password_input.send_keys(password)

    # Warte, bis der Anmelde-Button klickbar ist, und klicke darauf
    login_button = wait_for_element_to_be_clickable(driver, (By.ID, 'idSIButton9'))
    login_button.click()

    return extract_2fa_code_and_display(driver)


def open_time_booking_page_in_projektron(driver):
    projektron_domain = get_excel_loader().vba_settings_sheet.range(PROJEKTRON_DOMAIN_CELL).value.rstrip("/")
    full_projektron_url = f'{projektron_domain}/bcs/mybcs/dayeffortrecording/display'
    driver.get(f'{full_projektron_url}')
    return full_projektron_url


def extract_2fa_code_and_display(driver):
    # Wait for the 2FA number to be displayed
    element = WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, 'idRichContext_DisplaySign'))
    )
    return element.text  # Extract the 2FA number


def close_popups_in_projektron(driver):
    # Finde den Button für die Benachrichtigungen ("Ja" klicken)
    # Warte hier 20 Sekunden, da der vorherige Schritt die 2FA ist und man schon etwas Zeit haben sollte um sie durchzuführen
    notification_button = wait_for_element_to_be_clickable(driver, (By.CLASS_NAME, 'notificationPermissionLater'), 20)
    notification_button.click()

    # Warte auf den "Schließen"-Button mit einem genaueren XPath-Selektor
    close_button = wait_for_element_to_be_clickable(driver, (By.XPATH, '//div[@id="neutral_0"]//a[@class="close"]'))

    # Klicke auf den Schließen-Button
    close_button.click()


def find_terminal_notifier_path():
    terminal_notifier_path = shutil.which('terminal-notifier')
    if not terminal_notifier_path:
        # Fallback auf bekannten Standardpfad (falls vorhanden)
        possible_paths = [
            '/opt/homebrew/bin/terminal-notifier',  # Homebrew (macOS ARM64)
            '/usr/local/bin/terminal-notifier',    # Homebrew (Intel macOS)
            '/usr/bin/terminal-notifier'           # Standardpfade auf Linux
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
    return terminal_notifier_path


def display2FACode(two_fa_code):
    terminal_notifier_path = find_terminal_notifier_path()
    if not terminal_notifier_path:
        raise FileNotFoundError("Der Terminal-Notifier wurde nicht gefunden. Bitte installieren Sie sie.")
    subprocess.run([terminal_notifier_path, '-title', "2FA Code", '-message', f"Your 2FA Code is: {two_fa_code}"])


def clickSignInButton(driver):
    # "Ja" klicken, wenn gefragt wird, ob du angemeldet bleiben möchtest
    stay_signed_in_button = wait_for_optional_element_to_be_clickable(driver, (By.ID, 'idSIButton9'), duration=2)
    if stay_signed_in_button:
        stay_signed_in_button.click()


def fetch_projektron_task_main(sheet_name):
    excel_loader = ExcelLoader()
    time_sheet, vba_settings_sheet, wb = excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)
    response = ''

    headless = get_headless_mode(vba_settings_sheet)

    password_reference = get_excel_loader().vba_settings_sheet.range(ONE_PASSWORD_REFERENCE_PROJEKTRON_CELL).value
    projektron_credentials = get_credentials(sheet_name, password_reference)
    user = projektron_credentials['username']
    password = projektron_credentials['password']

    if time_sheet is not None:
        get_excel_loader().log_to_excel("Setup Connection ...")
    # Hauptteil des Skripts
    if (headless == 'true'):
        get_excel_loader().log_to_excel("HeadLess Mode")
        chrome_options = Options()
        chrome_options.headless = True
        chrome_options.add_argument("--headless")
        driver = webdriver.Chrome(options=chrome_options)
    else:
        driver = webdriver.Chrome()

    output_queue = queue.Queue()
    projektronLogin(driver, password, user, output_queue)

    # Warten auf die Ergebnisse aus der Queue
    driver, two_fa_code = output_queue.get()  # Hier erhalten Sie die Werte zurück

    # display 2fa Code in pupup
    display2FACode(two_fa_code)

    clickSignInButton(driver)

    close_popups_in_projektron(driver)

    start_spinner(driver)

    projektron_user_locale = get_excel_loader().vba_settings_sheet.range(PROJETRON_LOCALE_SETTING_CELL).value
    is_locale_set = projektron_user_locale != None and projektron_user_locale != ""

    if not is_locale_set:
        projektron_user_locale = get_user_language(driver)
        get_excel_loader().vba_settings_sheet.range(PROJETRON_LOCALE_SETTING_CELL).value = projektron_user_locale
        # Open Task Booking Page again, to proceed booking
        open_time_booking_page_in_projektron(driver)

    get_excel_loader().log_to_excel("Fetching Projektron Tasks Information...")
    sync_projektron_tasks(driver)

    stop_spinner(driver)

    return response


def get_headless_mode(settings_sheet):
    return 'true' if settings_sheet.range(HEADLESS_MODE_CELL).value == 'false' else 'false'


def projektronLogin(driver, password, user, output_queue):
    # Optionen für den Chrome-Browser erstellen
    two_fa_code = ''

    try:
        two_fa_code = login_to_website(driver, user, password)

    except Exception as e:
        get_excel_loader().log_to_excel(f"Ein Fehler ist aufgetreten: {e}")
        driver.quit()

    output_queue.put((driver, two_fa_code))


def get_user_language(driver):
    projektron_domain = get_excel_loader().vba_settings_sheet.range(PROJEKTRON_DOMAIN_CELL).value.rstrip("/")
    """
    Überprüft die Spracheinstellung des Benutzers im Profil.
    """
    # Gehe zu den Profileinstellungen
    profile_url = f'{projektron_domain}/bcs/mybcs/profile/edit'

    driver.get(profile_url)

    # Finde das Dropdown-Element für die Sprache
    language_select_element = driver.find_element(By.ID, "label_default_lang_lang")
    language_dropdown = Select(language_select_element)

    # Hole die aktuell ausgewählte Sprache
    selected_language = language_dropdown.first_selected_option.get_attribute("value")

    return selected_language


def fetch_projektron_tasks(driver):
    # Alle Zeilen finden, die Aufgaben enthalten
    task_rows = driver.find_elements(By.CSS_SELECTOR, 'tr.row.dragVisualisationTarget.default.selectableRow')

    tasks = []
    for row in task_rows:
        # Extrahiere die Task-ID
        task_id = row.get_attribute("data-taskoid")

        # Extrahiere die Beschreibungen
        description_elements = row.find_elements(By.CSS_SELECTOR, "td.content.blueMarkRow span.hover")
        descriptions = [elem.text.strip() for elem in description_elements]

        # Umgekehrte Reihenfolge der Beschreibungen
        reversed_descriptions = descriptions[::-1]  # Liste umdrehen

        # Kombiniere Beschreibungen zu einem String
        full_description = " ### ".join(reversed_descriptions)

        # Füge die Aufgabe zur Liste hinzu
        if task_id and full_description:
            tasks.append({"key": task_id, "description": full_description})

    return tasks


def write_tasks_to_excel(tasks):
    # Startzeile für Excel
    start_row = 4
    end_row = 54

    # Sicherstellen, dass die Daten in den Bereich passen
    if len(tasks) > (end_row - start_row + 1):
        get_excel_loader().log_to_excel(f"Zu viele Aufgaben ({len(tasks)}), um in den Bereich A{start_row}:B{end_row} zu passen.")
        tasks = tasks[: (end_row - start_row + 1)]  # Truncate auf maximalen Bereich

    # Schreiben der Daten in die Excel-Zellen
    for i, task in enumerate(tasks):
        row_index = start_row + i
        get_excel_loader().vba_settings_sheet.range(f"A{row_index}").value = task["key"]
        get_excel_loader().vba_settings_sheet.range(f"B{row_index}").value = task["description"]

    get_excel_loader().log_to_excel(f"{len(tasks)} Aufgaben wurden erfolgreich in Excel geschrieben.")


# Methode zum Starten des Spinners
def start_spinner(driver):
    loading_animation = """
    // Styles für das Overlay
    var style = document.createElement('style');
    style.type = 'text/css';
    style.innerHTML = `
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    #loading-overlay {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-color: rgba(0, 0, 0, 0.7);
        z-index: 9999;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        font-family: Arial, sans-serif;
        color: white;
    }
    .spinner {
        border: 10px solid #f3f3f3;
        border-top: 10px solid #3498db;
        border-radius: 50%;
        width: 80px;
        height: 80px;
        animation: spin 1s linear infinite;
    }
    #loading-text {
        margin-top: 20px;
        font-size: 20px;
        text-align: center;
        background: none; /* Entferne den Hintergrund der Box */
        padding: 0;      /* Entferne Innenabstand */
        color: white;    /* Textfarbe bleibt weiß */
    }
    `;
    document.head.appendChild(style);

    // Overlay erstellen
    var loadingDiv = document.createElement('div');
    loadingDiv.id = 'loading-overlay';

    // Spinner erstellen
    var spinner = document.createElement('div');
    spinner.className = 'spinner';

    // Text erstellen
    var loadingText = document.createElement('div');
    loadingText.id = 'loading-text';
    loadingText.innerHTML = 'Übertrage Projektron Tasks nach Excel... Bitte warten.';

    // Elemente zusammenfügen
    loadingDiv.appendChild(spinner);
    loadingDiv.appendChild(loadingText);
    document.body.appendChild(loadingDiv);
    """
    driver.execute_script(loading_animation)

# Methode zum Stoppen des Spinners
def stop_spinner(driver):
    remove_overlay = """
    var overlay = document.getElementById('loading-overlay');
    if (overlay) {
        overlay.remove();
    }
    """
    driver.execute_script(remove_overlay)


def sync_projektron_tasks(driver):
    try:
        # Aufgaben aus Projektron extrahieren
        tasks = fetch_projektron_tasks(driver)

        if not tasks:
            get_excel_loader().log_to_excel("Keine Aufgaben gefunden.")
            return

        # Aufgaben in Excel schreiben
        write_tasks_to_excel(tasks)

        get_excel_loader().log_to_excel("Projektron-Aufgaben erfolgreich synchronisiert.")
    except Exception as e:
        get_excel_loader().log_to_excel(f"Fehler beim Synchronisieren der Projektron-Aufgaben: {str(e)}")



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
