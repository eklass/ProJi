# coding: cp1252
import sys
import traceback
import requests
from jira import checkJiraTimes
from utils.getPasswordFrom1Password import get_credentials

from utils.Constants import JIRA_TICKET_COLUMN, TIME_COLUMN, COMMENT_COLUMN, BOOK_COMMENTS_TO_JIRA_CHECK_CELL, \
    JIRA_DOMAIN_CELL, ONE_PASSWORD_REFERENCE_JIRA_CELL
from utils.excelLoader import ExcelLoader, extract_time_from_cell, format_duration

global_excel_loader = None


def is_book_comments_to_jira_active():
    return 'true' if get_excel_loader().vba_settings_sheet.range(BOOK_COMMENTS_TO_JIRA_CHECK_CELL).value == 'true' else 'false'


def post_worklog_to_jira(ticket_number, duration_formatted, comment, date, headers):
    jira_domain = get_excel_loader().vba_settings_sheet.range(JIRA_DOMAIN_CELL).value.rstrip("/")
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
def post_jira_times(sheet_name):
    excel_loader = ExcelLoader()
    time_sheet, vba_settings_sheet, wb = excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)

    password_reference = get_excel_loader().vba_settings_sheet.range(ONE_PASSWORD_REFERENCE_JIRA_CELL).value

    try:
        get_excel_loader().log_to_excel("Posting Jira Times... ")

        jira_credentials = get_credentials(sheet_name, password_reference)
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
        user_info = get_user_info(headers)
        user_locale = user_info.get("locale")

        for row in range(7, 22):
            ticket_number = time_sheet.range(f'{JIRA_TICKET_COLUMN}{row}').value
            duration = time_sheet.range(f'{TIME_COLUMN}{row}').value

            duration_formatted = extract_time_from_cell(time_sheet.range(f'{TIME_COLUMN}{row}')) if isinstance(duration, float) else duration
            if not duration_formatted or duration_formatted == 0.0:
                continue

            duration_with_correct_locale = format_duration(duration_formatted, user_locale)
            if ticket_number is not None:
                if not checkJiraTimes.compare_jira_and_excel_times(sheet_name, row, date):
                    get_excel_loader().log_to_excel(f"###### {date.strftime('%Y-%m-%d')} ######\n")
                    comment = time_sheet.range(f'{COMMENT_COLUMN}{row}').value
                    response = post_worklog_to_jira(ticket_number, duration_with_correct_locale, comment, date, headers)
                    if response.status_code == 201:
                        get_excel_loader().log_to_excel(f"Arbeitszeit fuer Ticket {ticket_number} erfolgreich zurueckgemeldet.\n")
                    else:
                        get_excel_loader().log_to_excel(f"Fehler beim Zurueckmelden der Arbeitszeit fuer Ticket {ticket_number}: {response.text}\n")
                else:
                    get_excel_loader().log_to_excel(f"Fuer Ticket {ticket_number} gibt es bereits eine passende Zeit.")

        test_jira_output = checkJiraTimes.check_jira_times(sheet_name)
        if test_jira_output is not None:
            get_excel_loader().log_to_excel(test_jira_output + '\n')
        else:
            get_excel_loader().log_to_excel("Schaut gut aus ;)")

        get_open_tickets_for_user(user_info.get("email"), headers)

        wb.save()


    except Exception as e:
        stacktrace = traceback.format_exc()  # Stacktrace als String abrufen
        get_excel_loader().log_to_excel("Fehler beim Ausführen von post_jira_times: " + str(e) + "\n" + stacktrace)


def get_user_info(headers):
    jira_domain = get_excel_loader().vba_settings_sheet.range(JIRA_DOMAIN_CELL).value.rstrip("/")
    url = f"{jira_domain}/rest/api/2/myself"

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        user_data = response.json()
        # Extrahiere Locale und E-Mail
        user_locale = user_data.get("locale", "en_US")  # Fallback auf 'en_US'
        user_email = user_data.get("emailAddress", None)  # Fallback auf None, falls nicht vorhanden
        return {"locale": user_locale, "email": user_email}
    else:
        print(f"Fehler beim Abrufen der Benutzerdaten: {response.status_code}")
        return None


def get_open_tickets_for_user(user_email, headers):
    jira_domain = get_excel_loader().vba_settings_sheet.range(JIRA_DOMAIN_CELL).value.rstrip("/")
    # JQL-Query, um offene Tickets eines Benutzers zu finden
    jql_query = f'assignee = "{user_email}" AND statusCategory != Done ORDER BY created DESC'
    url = f'{jira_domain}/rest/api/2/search'

    # API-Request mit der JQL-Abfrage
    params = {"jql": jql_query, "fields": "key"}  # Wir benötigen zunächst nur die Ticketnummer
    response = requests.get(url, headers=headers, params=params)

    if response.status_code == 200:
        tickets = response.json().get("issues", [])
        ticket_keys = [ticket["key"] for ticket in tickets]
        return ticket_keys
    else:
        raise Exception(f"Fehler beim Abrufen der Tickets: {response.status_code} - {response.text}")



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
    post_jira_times("Dienstag")


if __name__ == "__main__":
    output = main()
    print(output)
