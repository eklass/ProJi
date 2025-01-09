# coding: cp1252
import sys
import traceback
import requests
import checkJiraTimes
import getPasswordFrom1Password
from excelLoader import ExcelLoader
import locale

jira_ticket_column = 'B'
time_column = 'D'
comment_column = 'C'
jira_status_column = 'K'
console_output_cell = 'E27'
vba_settings_sheet_name = 'VBA-Settings'
jira_domain_cell = 'H13'
book_comments_to_jira_check_cell = 'H14'
global_excel_loader = None

def extract_time_from_cell(cell):
    cell_value = cell.value
    try:
        # Versuche, den Wert in einen float zu konvertieren
        hours = float(cell_value) * 24  # Zelle wird in Stunden umgerechnet
        return hours  # Gib die Stunden als Float zurück
    except (ValueError, TypeError):
        # Gib eine sinnvolle Fehlermeldung oder einen Fallback zurück, falls die Konvertierung fehlschlägt
        return None


def is_book_comments_to_jira_active():
    global book_comments_to_jira_check_cell
    return 'true' if get_excel_loader().vba_settings_sheet.range(book_comments_to_jira_check_cell).value == 'true' else 'false'


def post_worklog_to_jira(ticket_number, duration_formatted, comment, date, headers):
    global jira_domain_cell

    jira_domain = get_excel_loader().vba_settings_sheet.range(jira_domain_cell).value.rstrip("/")
    full_jira_url = f'{jira_domain}/rest/api/2/issue/{{issueKey}}/worklog'

    if is_book_comments_to_jira_active() == 'false' or comment is None:
        comment = ''

    payload = {
        'timeSpent': duration_formatted,
        'comment': comment,
        'started': date.strftime('%Y-%m-%d') + 'T00:00:00.000+0000'
    }

    # Convert payload to a JSON string
    #payload_str = json.dumps(payload, indent=4)

    # Log the payload to Excel
    #log_to_excel(f"Payload: {payload_str}")

    try:
        response = requests.post(full_jira_url.format(issueKey=ticket_number), json=payload, headers=headers)
        return response
    except Exception as e:
        get_excel_loader().log_to_excel("Fehler beim Ausführen von post_worklog_to_jira.  " + str(e))
        sys.exit(1)


# Method will be called from Excel Makro via xlwings
def post_jira_times(sheet_name, password_reference):
    excel_loader = ExcelLoader()
    time_sheet, vba_settings_sheet, wb = excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)

    try:
        get_excel_loader().log_to_excel("Posting Jira Times... ")

        jira_credentials = getPasswordFrom1Password.get_credentials(sheet_name, password_reference)
        personal_access_token = jira_credentials['password']
        headers = {
            'Authorization': f'Bearer {personal_access_token}'
        }

        if time_sheet is None:
            get_excel_loader().log_to_excel(f"Das Blatt '{sheet_name}' wurde nicht gefunden.\n")
            return

        date = time_sheet.range('B1').value
        if date is None:
            get_excel_loader().log_to_excel("Ungueltiges Datum. Bitte ueberpruefen Sie das Datum.\n")
            return

        checkJiraTimes.set_headers(sheet_name, password_reference)

        for row in range(7, 22):
            ticket_number = time_sheet.range(f'{jira_ticket_column}{row}').value
            duration = time_sheet.range(f'{time_column}{row}').value
            user_locale = get_user_locale(headers)

            duration_formatted = extract_time_from_cell(time_sheet.range(f'{time_column}{row}')) if isinstance(duration, float) else duration
            if not duration_formatted or duration_formatted == 0.0:
                continue

            duration_with_correct_locale = format_duration(duration_formatted, user_locale)
            if ticket_number is not None:
                if not checkJiraTimes.compare_jira_and_excel_times(sheet_name, row, date):
                    get_excel_loader().log_to_excel(f"###### {date.strftime('%Y-%m-%d')} ######\n")
                    comment = time_sheet.range(f'{comment_column}{row}').value
                    response = post_worklog_to_jira(ticket_number, duration_with_correct_locale, comment, date, headers)
                    if response.status_code == 201:
                        get_excel_loader().log_to_excel(f"Arbeitszeit fuer Ticket {ticket_number} erfolgreich zurueckgemeldet.\n")
                    else:
                        get_excel_loader().log_to_excel(f"Fehler beim Zurueckmelden der Arbeitszeit fuer Ticket {ticket_number}: {response.text}\n")
                else:
                    get_excel_loader().log_to_excel(f"Fuer Ticket {ticket_number} gibt es bereits eine passende Zeit.")

        test_jira_output = checkJiraTimes.check_jira_times(sheet_name, password_reference)
        if test_jira_output is not None:
            get_excel_loader().log_to_excel(test_jira_output + '\n')
        else:
            get_excel_loader().log_to_excel("Schaut gut aus ;)")

        wb.save()


    except Exception as e:
        stacktrace = traceback.format_exc()  # Stacktrace als String abrufen
        get_excel_loader().log_to_excel("Fehler beim Ausführen von post_jira_times: " + str(e) + "\n" + stacktrace)


def get_user_locale(headers):
    jira_domain = get_excel_loader().vba_settings_sheet.range(jira_domain_cell).value.rstrip("/")
    url = f"{jira_domain}/rest/api/2/myself"

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        user_data = response.json()
        return user_data.get("locale", "en_US")  # Fallback auf 'en_US', falls nicht vorhanden
    else:
        print(f"Fehler beim Abrufen der Benutzerdaten: {response.status_code}")
        return None


def format_duration(duration_hours, user_locale):
    locale.setlocale(locale.LC_ALL, user_locale)  # Beispiel für Deutsch
    return locale.format_string("%.2f", duration_hours)


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

def main():
    post_jira_times("Dienstag", "JIRA-API-KEY")


if __name__ == "__main__":
    output = main()
    print(output)
