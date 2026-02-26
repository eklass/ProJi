# coding: cp1252
import sys
import traceback
import requests
from utils.JiraUtils import (
    get_jira_mapping_for_ticket_number,
    get_password_reference,
    get_jira_domain,
    create_http_header,
    get_worklog_url_for_ticket_number
)
from jira.checkJiraTimes import fetch_jira_data, check_jira_times
from utils.Constants import (
    WEEKDAY_JIRA_TICKET_COLUMN,
    WEEKDAY_TIME_COLUMN,
    WEEKDAY_COMMENT_COLUMN,
    BOOK_COMMENTS_TO_JIRA_CHECK_CELL,
    PROJI_SETTINGS_JIRA_HOST_COLUMN
)
from utils.excelLoader import ExcelLoader, extract_time_from_cell, format_duration, convert_time_to_decimal

global_excel_loader = None

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

def is_book_comments_to_jira_active():
    return 'true' if get_excel_loader().vba_settings_sheet.range(BOOK_COMMENTS_TO_JIRA_CHECK_CELL).value == 'true' else 'false'

def post_worklog_to_jira(session, jira_domain, ticket_number, duration_formatted, comment, date):
    if is_book_comments_to_jira_active() == 'false' or not comment:
        comment = ''
    payload = {
        'timeSpent': duration_formatted,
        'comment': comment,
        'started': date.strftime('%Y-%m-%d') + 'T00:00:00.000+0000'
    }
    url = f"{jira_domain}/rest/api/2/issue/{ticket_number}/worklog"
    try:
        return session.post(url, json=payload)
    except Exception as e:
        get_excel_loader().log_to_excel(f"Fehler beim Ausführen von post_worklog_to_jira: {e}")
        sys.exit(1)

def get_user_info(headers, jira_domain):
    jira_domain = jira_domain.rstrip("/")
    url = f"{jira_domain}/rest/api/2/myself"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        data = response.json()
        return {"locale": data.get("locale", "en_US"), "email": data.get("emailAddress")}
    else:
        get_excel_loader().log_to_excel(f"Fehler beim Abrufen der Benutzerdaten: {response.status_code}")
        return {"locale": "en_US", "email": None}

def post_jira_times(sheet_name):
    excel_loader = ExcelLoader()
    time_sheet, vba_settings_sheet, wb = excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)

    try:
        get_excel_loader().log_to_excel("Posting Jira Times... ")
        if time_sheet is None:
            get_excel_loader().log_to_excel(f"Das Blatt '{sheet_name}' wurde nicht gefunden.\n")
            return

        date = time_sheet.range('B1').value
        if date is None:
            get_excel_loader().log_to_excel("Ungueltiges Datum. Bitte ueberpruefen Sie das Datum.\n")
            return

        # Excel-Bereiche in einem Batch
        tickets   = time_sheet.range(f'{WEEKDAY_JIRA_TICKET_COLUMN}7:{WEEKDAY_JIRA_TICKET_COLUMN}21').value
        durations = time_sheet.range(f'{WEEKDAY_TIME_COLUMN}7:{WEEKDAY_TIME_COLUMN}21').value
        comments  = time_sheet.range(f'{WEEKDAY_COMMENT_COLUMN}7:{WEEKDAY_COMMENT_COLUMN}21').value

        session_cache = {}
        proj_sheet = get_excel_loader().get_sheet("ProJi-Settings")

        for idx, (ticket_number, duration, comment) in enumerate(zip(tickets, durations, comments), start=7):
            if not ticket_number or not duration:
                continue

            # Format Excel-Dauer
            if isinstance(duration, float):
                duration_formatted = extract_time_from_cell(time_sheet.range(f'{WEEKDAY_TIME_COLUMN}{idx}'))
            else:
                duration_formatted = duration
            if not duration_formatted or duration_formatted == 0.0:
                continue

            # Jira-Mapping
            mapping_row = get_jira_mapping_for_ticket_number(ticket_number, proj_sheet)
            password_ref  = get_password_reference(mapping_row)
            jira_domain   = get_jira_domain(mapping_row).rstrip("/")
            worklog_url   = get_worklog_url_for_ticket_number(ticket_number, mapping_row)

            # Session pro Domain
            session = session_cache.get(jira_domain)
            if session is None:
                session = requests.Session()
                headers = create_http_header(sheet_name, password_ref)
                session.headers.update(headers)
                session_cache[jira_domain] = session

            # User Info für Locale
            user_info = get_user_info(session.headers, jira_domain)
            duration_localized = format_duration(duration_formatted, user_info.get("locale"))

            # Jira-Daten holen und vergleichen
            date_str = date.strftime('%Y-%m-%d')
            jira_times = fetch_jira_data(session, worklog_url, date_str, comment or '', user_info.get('email'))
            jira_decimal = [convert_time_to_decimal(t) for t in jira_times]

            if duration_formatted not in jira_decimal:
                get_excel_loader().log_to_excel(f"###### {date.strftime('%Y-%m-%d')} ######\n")
                response = post_worklog_to_jira(
                    session, jira_domain, ticket_number, duration_localized, comment, date
                )
                if response.status_code == 201:
                    get_excel_loader().log_to_excel(f"Arbeitszeit fuer Ticket {ticket_number} erfolgreich zurueckgemeldet.\n")
                else:
                    get_excel_loader().log_to_excel(
                        f"Fehler beim Zurueckmelden fuer Ticket {ticket_number}: {response.text}\n"
                    )
            else:
                get_excel_loader().log_to_excel(f"Fuer Ticket {ticket_number} gibt es bereits eine passende Zeit.")

        # Abschließender Check
        result = check_jira_times(sheet_name)
        get_excel_loader().log_to_excel((result or "Schaut gut aus ;)") + "\n")

        wb.save()

    except Exception as e:
        stack = traceback.format_exc()
        get_excel_loader().log_to_excel(f"Fehler beim Ausführen von post_jira_times: {e}\n{stack}")

if __name__ == "__main__":
    post_jira_times("Mittwoch")
