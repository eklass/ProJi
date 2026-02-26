# coding: cp1252
import sys
import traceback
import requests
from utils.JiraUtils import (
    create_http_header,
    get_jira_mapping_for_ticket_number,
    get_password_reference,
    get_worklog_url_for_ticket_number
)
from utils.Constants import (
    EMAIL_CELL,
    WEEKDAY_JIRA_TICKET_COLUMN,
    WEEKDAY_TIME_COLUMN,
    WEEKDAY_JIRA_STATUS_COLUMN,
    WEEKDAY_COMMENT_COLUMN
)
from utils.excelLoader import ExcelLoader, extract_time_from_cell, convert_time_to_decimal

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

def fetch_jira_data(session, worklog_url, date_str, comment, user_email):
    """
    Ruft Worklogs für ein einzelnes Ticket ab und filtert nach Autor, Datum und Kommentar.
    """
    try:
        response = session.get(worklog_url)
    except Exception as e:
        get_excel_loader().log_to_excel(f"Fehler beim Abrufen von Jira-Daten ({worklog_url}): {e}")
        return []

    if response.status_code != 200:
        get_excel_loader().log_to_excel(
            f"Fehler im Jira-Response für {worklog_url}: {response.status_code} {response.text}"
        )
        return []

    worklogs = response.json().get('worklogs', []) or []
    matching = []
    for wl in worklogs:
        if (wl.get('author', {}).get('emailAddress') == user_email and
                wl.get('started', '').startswith(date_str) and
                wl.get('comment', '') == (comment or '')):
            secs = int(wl.get('timeSpentSeconds', 0))
            hours = secs // 3600
            mins = (secs % 3600) // 60
            matching.append(f"{hours:02d}:{mins:02d}")
    return matching

def check_jira_times(sheet_name):
    excel_loader = ExcelLoader()
    time_sheet, vba_settings_sheet, wb = excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)

    try:
        get_excel_loader().log_to_excel("Jira check ...")
        if not time_sheet:
            get_excel_loader().log_to_excel(f"Das Blatt '{sheet_name}' wurde nicht gefunden.")
            return

        date = time_sheet.range('B1').value
        if not date:
            get_excel_loader().log_to_excel("Ungültiges Datum. Bitte überprüfen Sie das Datum.")
            return
        date_str = date.strftime('%Y-%m-%d')

        # Nutzer-E-Mail
        user_email = vba_settings_sheet.range(EMAIL_CELL).value or ''

        # Projekt-Mapping einmalig
        proj_sheet = get_excel_loader().get_sheet("ProJi-Settings")

        # Bereiche in einem Batch laden
        tickets  = time_sheet.range(f'{WEEKDAY_JIRA_TICKET_COLUMN}7:{WEEKDAY_JIRA_TICKET_COLUMN}21').value
        durations= time_sheet.range(f'{WEEKDAY_TIME_COLUMN}7:{WEEKDAY_TIME_COLUMN}21').value
        comments = time_sheet.range(f'{WEEKDAY_COMMENT_COLUMN}7:{WEEKDAY_COMMENT_COLUMN}21').value

        session_cache = {}
        found_any = False

        for idx, (ticket, duration, comment) in enumerate(zip(tickets, durations, comments), start=7):
            if not ticket:
                continue
            found_any = True

            # Jira-URL und Domain
            mapping_row = get_jira_mapping_for_ticket_number(ticket, proj_sheet)
            worklog_url = get_worklog_url_for_ticket_number(ticket, mapping_row)
            jira_domain = worklog_url.split('/rest/')[0].rstrip('/')

            # Session reuse pro Domain
            sess = session_cache.get(jira_domain)
            if not sess:
                sess = requests.Session()
                pw_ref = get_password_reference(mapping_row)
                headers = create_http_header(sheet_name, pw_ref)
                sess.headers.update(headers)
                session_cache[jira_domain] = sess

            # Dauer aus Excel
            if isinstance(duration, float):
                dur_fmt = extract_time_from_cell(time_sheet.range(f'{WEEKDAY_TIME_COLUMN}{idx}'))
            else:
                dur_fmt = duration or ''

            # Worklogs abrufen
            jira_times = fetch_jira_data(sess, worklog_url, date_str, comment or '', user_email)
            jira_decimal = [convert_time_to_decimal(t) for t in jira_times]

            # Vergleich und Status setzen
            if dur_fmt in jira_decimal:
                status = 'passt'
            else:
                status = 'passt nicht'
            time_sheet.range(f'{WEEKDAY_JIRA_STATUS_COLUMN}{idx}').value = status

            # Logging
            founds = ", ".join(jira_times) if jira_times else "keine"
            get_excel_loader().log_to_excel(
                f"Ticket {ticket}: Excel {dur_fmt}h -> Jira {founds} -> {status}\n"
            )

        if not found_any:
            get_excel_loader().log_to_excel("Kein Jira Ticket im aktuellen Sheet gefunden")

        wb.save()

    except Exception as e:
        stack = traceback.format_exc()
        get_excel_loader().log_to_excel(f"Fehler beim Ausführen von check_jira_times: {e}\n{stack}")

def main():
    check_jira_times("Montag")

if __name__ == "__main__":
    main()
