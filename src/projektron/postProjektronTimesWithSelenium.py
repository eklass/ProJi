import os
import queue
import shutil
import subprocess
import time
import traceback

from datetime import datetime
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService

from utils.getPasswordFrom1Password import get_credentials
from utils.Constants import PROJEKTRON_LOCALE_SETTING_CELL, WEEKDAY_PROJEKTRON_STATUS_COLUMN, PROJEKTRON_DOMAIN_CELL, \
    ONE_PASSWORD_REFERENCE_PROJEKTRON_CELL
from utils.excelLoader import ExcelLoader, format_duration

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

    # Überprüfe, ob das Passwortfeld vorhanden ist
    try:
        password_input = wait_for_element_to_be_clickable(driver, (By.NAME, 'passwd'))
        password_input.send_keys(password)

        # Warte, bis der Anmelde-Button klickbar ist, und klicke darauf
        login_button = wait_for_element_to_be_clickable(driver, (By.ID, 'idSIButton9'))
        login_button.click()

    except TimeoutException:
        # Falls das Passwortfeld nicht erscheint, fahre direkt fort
        get_excel_loader().log_to_excel("Passwortfeld wurde nicht angezeigt, weiter ohne Eingabe.")

    return extract_2fa_code_and_display(driver)

def open_time_booking_page_in_projektron(driver):
    projektron_domain = get_excel_loader().vba_settings_sheet.range(PROJEKTRON_DOMAIN_CELL).value.rstrip("/")
    full_projektron_url = f'{projektron_domain}/bcs/mybcs/dayeffortrecording/display'
    driver.get(f'{full_projektron_url}')
    return full_projektron_url

def extract_2fa_code_and_display(driver):
    try:
        # Warte, bis das 2FA-Element sichtbar ist
        element = WebDriverWait(driver, 3).until(
            EC.visibility_of_element_located((By.ID, 'idRichContext_DisplaySign'))
        )
        return element.text  # Extrahiere und gib den 2FA-Code zurück
    except Exception:
        # Falls das 2FA-Element nicht gefunden wird, informiere über das Fehlen und gib None zurück
        get_excel_loader().log_to_excel("2FA-Code wurde nicht angezeigt, weiter ohne Code.")
        return None


def close_popups_in_projektron(driver):
    # Finde den Button für die Benachrichtigungen ("Ja" klicken)
    # Warte hier 20 Sekunden, da der vorherige Schritt die 2FA ist und man schon etwas Zeit haben sollte um sie durchzuführen
    notification_button = wait_for_element_to_be_clickable(driver, (By.CLASS_NAME, 'notificationPermissionLater'), 20)
    notification_button.click()

    # Warte auf den "Schließen"-Button mit einem genaueren XPath-Selektor
    close_button = wait_for_element_to_be_clickable(driver, (By.XPATH, '//div[@id="neutral_0"]//a[@class="close"]'))

    # Klicke auf den Schließen-Button
    close_button.click()


def select_date(driver, day, month, year):
    date_display = wait_for_element_to_be_clickable(driver, (By.ID, "daytimerecording,Selections,effortRecordingDate_intervaldisplay"))
    date_display.click()
    day_link = wait_for_element_to_be_clickable(driver, (By.XPATH, f"//a[text()='{day}' and contains(@href, 'year={year}') and contains(@href, 'month={month}')]"))
    day_link.click()


def create_and_fill_tasks(driver, date_day, date_month, date_year, task_details_list, projektron_user_locale):
    select_date(driver, date_day, date_month, date_year)
    for task_details in task_details_list:
        task_group_oid = task_details['task_group_oid']
        description = task_details['description']
        row_in_timesheet = task_details['row_in_timesheet']
        duration = task_details['duration']
        formatted_hours = format_duration(duration, projektron_user_locale)

        try:
            task_row = add_task_row(driver, task_group_oid)
            fill_task_details(task_row, formatted_hours, description)
        except TimeoutException:
            get_excel_loader().log_to_excel(f"TimeoutException: TaskID {task_group_oid} konnte nicht gefunden werden.")
            get_excel_loader().get_time_sheet().range(WEEKDAY_PROJEKTRON_STATUS_COLUMN + row_in_timesheet).value = "Fehler"
            continue


def add_task_row(driver, task_group_oid):
    try:
        original_task_row = wait_for_element_to_be_clickable(driver, (By.CSS_SELECTOR, f'tr[data-listtaskgroupoid="{task_group_oid}"]'))
        duplicate_row_button = wait_for_element_to_be_clickable(original_task_row, (By.CSS_SELECTOR, "button[name*='duplicateEffortRow']"))

        # Klicke den Button per JavaScript
        driver.execute_script("arguments[0].click();", duplicate_row_button)

        # Nachdem die neue Zeile hinzugefügt wurde, aktualisiere die Liste der Task-Zeilen
        new_task_rows = driver.find_elements(By.CSS_SELECTOR, 'tr[data-listtaskgroupoid]')
        original_row_index = new_task_rows.index(original_task_row)
        return new_task_rows[original_row_index + 1]
    except TimeoutException as e:
        raise TimeoutException(f"TaskID {task_group_oid} nicht gefunden: {e}")


def fill_task_details(task_row, duration, description):
    duration_input = task_row.find_element(By.CSS_SELECTOR, "input[name*='effortExpense_hour']")
    duration_input.clear()
    duration_input.send_keys(duration)
    description_textarea = task_row.find_element(By.CSS_SELECTOR, "textarea[name*='description']")
    description_textarea.clear()
    description_textarea.send_keys(description)


def save(driver):
    # Warte, bis der Speichern-Button bereit ist zum Klicken
    save_button = wait_for_element_to_be_clickable(driver, (By.CSS_SELECTOR, "button.button.Apply.MainButton.possible_default_button.hasChanges"))

    # Klicke per JavaScript auf den Speichern-Button
    driver.execute_script("arguments[0].click();", save_button)
    time.sleep(2)
    

def task_exists(driver, description, duration, row_in_timesheet, message_in_case_of_missing_booking):

    # Suche nach allen Buchungszeilen
    booking_rows = driver.find_elements(By.CSS_SELECTOR, 'tr.row.dragVisualisationTarget.default.selectableRow')  # Der Selektor muss angepasst werden an deine Webseite

    for row in booking_rows:
        # Extrahiere die Beschreibung und Dauer
        task_description_elements = row.find_elements(By.CSS_SELECTOR, "td.content.blueMarkRow textarea.cellValueProvider")

        # Da es mehrere Textareas geben könnte, stellen wir sicher, dass wir Text aus allen extrahieren
        task_descriptions = [element.get_attribute('value') for element in task_description_elements if element.get_attribute('value').strip()]

        # Vergleiche die extrahierten Daten mit den gegebenen Parametern
        for task_descr in task_descriptions:
            if description == task_descr.strip():
                # and duration == task_duration:
                get_excel_loader().get_time_sheet().range(WEEKDAY_PROJEKTRON_STATUS_COLUMN + row_in_timesheet).value = "passt"
                return True

    get_excel_loader().get_time_sheet().range(WEEKDAY_PROJEKTRON_STATUS_COLUMN + row_in_timesheet).value = message_in_case_of_missing_booking
    return False


def filter_existing_tasks(driver, date_day, date_month, date_year, tasks_to_add, message_in_case_of_missing_booking):
    # Zuerst das Datum auswählen, um sicherzustellen, dass wir die richtigen Daten abrufen
    select_date(driver, date_day, date_month, date_year)

    return [
        task for task in tasks_to_add
        if not task_exists(driver, task['description'], task['duration'], task['row_in_timesheet'],message_in_case_of_missing_booking)
    ]


def get_and_print_response(driver, response):
    # Überprüfe, ob es eine Warnmeldung gibt
    time.sleep(2)
    warning_message_elements = driver.find_elements(By.CSS_SELECTOR, "div.msg.warning")
    for message_element in warning_message_elements:
        message_text = message_element.find_element(By.CSS_SELECTOR, "span").text
        if message_text:  # Überprüft, ob message_text nicht leer ist
            response += message_text
            print(f"Warnung: {message_text}")

    error_message_elements = driver.find_elements(By.CSS_SELECTOR, "div.msg.error")
    for message_element in error_message_elements:
        message_text = message_element.find_element(By.CSS_SELECTOR, "span").text
        if message_text:  # Überprüft, ob message_text nicht leer ist
            response += message_text
            print(f"Error: {message_text}")

    # Überprüfe, ob es eine Erfolgsmeldung gibt
    success_message_elements = driver.find_elements(By.ID, "TimeRecordingService_Success")
    for message_element in success_message_elements:
        message_text = message_element.find_element(By.CSS_SELECTOR, "span").text
        response += message_text
        print(f"Erfolg: {message_text}")

    # Wenn keine Meldungen gefunden wurden, gib eine Standardnachricht aus
    if not warning_message_elements and not success_message_elements and not error_message_elements:
        no_response_found = "Keine spezifischen Rückmeldungen gefunden."
        response += no_response_found
        print(no_response_found)

    return response


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


def main(tasks_to_add, date, sheet_name, headless):
    excel_loader = ExcelLoader()
    time_sheet, vba_settings_sheet, wb = excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)
    response = ''

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
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
    else:
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

    output_queue = queue.Queue()
    projektronLogin(driver, password, user, output_queue)

    # Warten auf die Ergebnisse aus der Queue
    driver, two_fa_code = output_queue.get()  # Hier erhalten Sie die Werte zurück

    # display 2fa Code in pupup
    display2FACode(two_fa_code)

    clickSignInButton(driver)

    close_popups_in_projektron(driver)

    projektron_user_locale = get_excel_loader().vba_settings_sheet.range(PROJEKTRON_LOCALE_SETTING_CELL).value
    is_locale_set = projektron_user_locale != None and projektron_user_locale != ""

    if not is_locale_set:
        projektron_user_locale = get_user_language(driver)
        get_excel_loader().vba_settings_sheet.range(PROJEKTRON_LOCALE_SETTING_CELL).value = projektron_user_locale
        # Open Task Booking Page again, to proceed booking
        open_time_booking_page_in_projektron(driver)


    # Filtere Tasks, die bereits existieren
    get_excel_loader().log_to_excel(f"Check Duplicate Booking for {len(tasks_to_add)} tasks ...")

    tasks_to_add = filter_existing_tasks(driver, date.day, date.month, date.year, tasks_to_add, "in progress")

    if tasks_to_add.__len__() > 0:
        # Füge nur neue Tasks hinzu
        get_excel_loader().log_to_excel(f"Booking {len(tasks_to_add)} tasks ...")
        create_and_fill_tasks(driver, date.day, date.month, date.year, tasks_to_add, projektron_user_locale)
        get_excel_loader().log_to_excel("Save ...")
        save(driver)  # Speichere alle gemachten Eingaben
        response = get_and_print_response(driver, response)  # Gibt die Rückmeldung der Seite in der Konsole aus
    else:
        no_entries_to_book = "No new entries to book!"
        response += no_entries_to_book
        print(no_entries_to_book)

    # Set Status
    filter_existing_tasks(driver, date.day, date.month, date.year, tasks_to_add, "passt nicht")

    return response


def projektronLogin(driver, password, user, output_queue):
    # Optionen für den Chrome-Browser erstellen
    two_fa_code = ''

    try:
        two_fa_code = login_to_website(driver, user, password)
    except Exception as e:
        # Fange den Fehler ab und logge ihn zusammen mit dem Stacktrace
        error_message = f"Ein Fehler ist aufgetreten: {e}\n{traceback.format_exc()}"
        get_excel_loader().log_to_excel(error_message)
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
    # Datenstruktur, die mehrere Tasks enthält
    tasks_to_add = [
        {'task_group_oid': '1650658262809_JTask', 'duration': 0.25, 'description': "Beschreibung für Task 1", 'row_in_timesheet' : '7'},
    ]
    output = main(tasks_to_add, datetime(2025, 1, 10), "Freitag", "false")
    print(output)
