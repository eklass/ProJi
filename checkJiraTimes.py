import sys
import traceback
import requests
import getPasswordFrom1Password
from excelLoader import ExcelLoader

jira_ticket_column = 'B'
time_column = 'D'
jira_status_column = 'K'
console_output_cell = 'E27'
session_token_cell = 'H5'
jira_domain_cell = 'H13'
email_cell = 'H4'


# Ausgabevariable für die gesammelten Nachrichten
output_messages = []
headers = None
ticket_found = None
global_excel_loader = None


# Funktion zum Abrufen der Arbeitsprotokolle für ein Ticket aus der Jira-API
def fetch_jira_data(ticket_number, date):
    global jira_domain_cell
    jira_domain = get_excel_loader().vba_settings_sheet.range(jira_domain_cell).value.rstrip("/")
    url_template = f'{jira_domain}/rest/api/2/issue/{{issue_key}}/worklog'

    if not ticket_number.strip():
        return None

    date_str = date.strftime('%Y-%m-%d')
    url = url_template.format(issue_key=ticket_number.strip())

    try:
        response = requests.get(url, headers=headers)
    except Exception as e:
        get_excel_loader().log_to_excel("Fehler beim Ausführen von fetch_jira_data: " + str(e))
        sys.exit(1)

    if response.status_code == 200:
        worklogs = response.json().get('worklogs', [])

        # Filter worklogs by author and date
        user_email = get_excel_loader().vba_settings_sheet.range(email_cell).value
        matching_durations = []

        for worklog in worklogs:
            if (worklog['author']['emailAddress'] == user_email
                    and worklog['started'].startswith(date_str)):

                total_time_logged_seconds = int(worklog['timeSpentSeconds'])
                total_time_logged_hours = total_time_logged_seconds // 3600
                total_time_logged_minutes = (total_time_logged_seconds % 3600) // 60

                matching_durations.append(f"{total_time_logged_hours:02d}:{total_time_logged_minutes:02d}")

        return matching_durations if matching_durations else None
    else:
        get_excel_loader().log_to_excel(f"Fehler beim Abrufen der Arbeitsprotokolle für Ticket {ticket_number}: {response.text}")
        return None



# Funktion zum Extrahieren der Zeit im Format "hh:mm"
def extract_time_from_cell(cell):
    cell_value = cell.value
    hours = int(cell_value * 24)
    minutes = int((cell_value * 24 * 60) % 60)
    return f"{hours:02}:{minutes:02}"


# Vergleich von Jira- und Excel-Zeiten
def compare_jira_and_excel_times(sheet_name, row, date):
    global ticket_found

    excel_loader = ExcelLoader()
    excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)

    time_sheet = get_excel_loader().get_sheet(sheet_name)
    ticket_number = time_sheet.range(f'{jira_ticket_column}{row}').value

    if ticket_number:
        ticket_found = True
        duration = time_sheet.range(f'{time_column}{row}').value
        duration_formatted = extract_time_from_cell(time_sheet.range(f'{time_column}{row}')) if isinstance(duration, float) else duration
        jira_data = fetch_jira_data(ticket_number, date)

        if jira_data and duration_formatted in jira_data:
            time_sheet.range(f'{jira_status_column}{row}').value = 'passt'
            get_excel_loader().log_to_excel(f"Jira vs Excel {ticket_number} passt")
            return True
        else:
            time_sheet.range(f'{jira_status_column}{row}').value = 'passt nicht'
            found_durations = ", ".join(jira_data) if jira_data else "keine"
            get_excel_loader().log_to_excel(f"In Jira hat das Ticket {ticket_number} keine passende Dauer gefunden. Excel erwartete: {duration_formatted} Stunden.\nGefundene Zeiten in Jira zum aktuellen Tag: {found_durations}.")
            return False


def check_jira_times(sheet_name, password_reference):
    global headers, ticket_found
    excel_loader = ExcelLoader()
    time_sheet, vba_settings_sheet, wb = excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)
    
    ticket_found = False
    try:
        get_excel_loader().log_to_excel("Jira check ...")
        set_headers(sheet_name, password_reference)
        if not time_sheet:
            get_excel_loader().log_to_excel(f"Das Blatt '{sheet_name}' wurde nicht gefunden.")
            return
        date = time_sheet.range('B1').value
        if not date:
            get_excel_loader().log_to_excel("Ungültiges Datum. Bitte überprüfen Sie das Datum.")
            return
        for row in range(7, 22):
            compare_jira_and_excel_times(sheet_name, row, date)
        if not ticket_found:
            get_excel_loader().log_to_excel("Kein Jira Ticket im aktuellen Sheet gefunden")

        wb.save()
    except Exception as e:
        stacktrace = traceback.format_exc()  # Stacktrace als String abrufen
        get_excel_loader().log_to_excel("Fehler beim Ausführen von check_jira_times: " + str(e) + "\n" + stacktrace)



def set_headers(sheet_name, password_reference):
    global headers

    # Hier setzen Sie die Zelle, die den session_token enthält
    jira_credentials = getPasswordFrom1Password.get_credentials(sheet_name, password_reference)
    personal_access_token = jira_credentials['password']
    if not personal_access_token:
        get_excel_loader().log_to_excel("Error during fetching personal accessToken from 1Password in checkJiraTimes")
        return

    headers = {'Authorization': f'Bearer {personal_access_token}'}


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


# Hauptfunktion zum Ausführen des Vergleichs
def main():
    check_jira_times("Dienstag", "JIRA-API-KEY")


if __name__ == "__main__":
    main()
