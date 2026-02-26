import requests

from utils.JiraUtils import get_password_reference, get_jira_mapping_for_ticket_number, get_jira_host_with_row, \
    get_worklog_url_for_ticket_number, create_http_header
from utils.getPasswordFrom1Password import get_credentials

from utils.Constants import VBA_SETTINGS_SHEET_NAME, VBA_SHEET_CONSOLE_OUTPUT_CELL, \
    WEEKDAY_JIRA_TICKET_DESCRIPTON_RANGE, WEEKDAY_JIRA_TICKET_NUMBER_RANGE, \
    AUTOCOMPLETE_JIRA_TICKET_DESCRIPTION_CELL, PRO_JI_SETTINGS
from utils.excelLoader import ExcelLoader

headers = None
global_excel_loader = None

def setup_excel(sheet_name):
    excel_loader = ExcelLoader(None, None, VBA_SHEET_CONSOLE_OUTPUT_CELL)
    excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)


def fetch_jira_ticket_information():
    setup_excel(VBA_SETTINGS_SHEET_NAME)
    proji_settings_sheet = get_excel_loader().get_sheet(PRO_JI_SETTINGS)

    get_excel_loader().log_to_excel("Started Fetching Jira Tickets for not finished Tasks assigned to this user ...", True)
    jira_hosts = get_jira_host_with_row(proji_settings_sheet)
    for jira_host, row in jira_hosts:
        password_reference = get_password_reference(row)
        headers = create_http_header(PRO_JI_SETTINGS, password_reference)

        # Save autocomplete flag and set it to false, to not call fetch_missing_descriptions_from_jira while executing this fetch_jira_ticket_information
        autocomplete_jira_ticket_description_before = get_excel_loader().vba_settings_sheet.range(AUTOCOMPLETE_JIRA_TICKET_DESCRIPTION_CELL).value
        get_excel_loader().vba_settings_sheet.range(AUTOCOMPLETE_JIRA_TICKET_DESCRIPTION_CELL).value = "false"

        set_headers(VBA_SETTINGS_SHEET_NAME, password_reference)

        # TODO: Iterate over Proji Settings Sheet and enrich tickets variable
        user_info = get_user_info(jira_host, headers)

        tickets = get_open_tickets_for_user(jira_host, user_info.get("email"), headers)
        sync_tickets_to_excel(tickets)
    get_excel_loader().log_to_excel(f"{len(tickets)} Tickets fetched from Jira")

    # restore setting
    get_excel_loader().vba_settings_sheet.range(AUTOCOMPLETE_JIRA_TICKET_DESCRIPTION_CELL).value = autocomplete_jira_ticket_description_before


def sync_tickets_to_excel(ticket_details):
    # Excel-Bereich für die Jira-Tickets
    description_range = get_excel_loader().vba_settings_sheet.range("D3:D53")
    ticket_number_range = get_excel_loader().vba_settings_sheet.range("E3:E53")

    # Bestehende Daten auslesen
    existing_descriptions = [cell.value for cell in description_range]
    existing_ticket_numbers = [cell.value for cell in ticket_number_range]

    # Liste der vorhandenen Tickets als Set für schnellen Abgleich
    existing_tickets = set(existing_ticket_numbers)

    # Startindex für das Schreiben (erste freie Zeile)
    start_index = len([val for val in existing_ticket_numbers if val]) + 3  # Suche nach erster freien Zeile (Index in Excel beginnt bei 3)

    for ticket in ticket_details:
        ticket_number = ticket["key"]
        description = ticket["key"] + " - " + ticket["description"]

        if ticket_number not in existing_tickets:
            if start_index <= 53:  # Sicherstellen, dass wir im Bereich D3:E53 bleiben
                # Schreibe Ticketnummer und Beschreibung in Excel
                get_excel_loader().vba_settings_sheet.range(f"D{start_index}").value = description
                get_excel_loader().vba_settings_sheet.range(f"E{start_index}").value = ticket_number
                start_index += 1
            else:
                get_excel_loader().log_to_excel("Excel-Bereich D3:E53 ist voll. Einige Tickets konnten nicht hinzugefügt werden.")
                break

    get_excel_loader().log_to_excel("Fehlende Tickets wurden ergänzt.")


def get_user_info(jira_domain, headers):
    url = f"{jira_domain}/rest/api/2/myself"

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        user_data = response.json()
        # Extrahiere Locale und E-Mail
        user_locale = user_data.get("locale", "en_US")  # Fallback auf 'en_US'
        user_email = user_data.get("emailAddress", None)  # Fallback auf None, falls nicht vorhanden
        return {"locale": user_locale, "email": user_email}
    else:
        get_excel_loader().log_to_excel(f"Fehler beim Abrufen der Benutzerdaten: {response.status_code}")
        return None


def log_tickets_to_excel(tickets):
    for ticket in tickets:
        # Schreibe Ticketnummer und Beschreibung in Excel
        get_excel_loader().log_to_excel(f"{ticket['key']} - {ticket['description']}")


def get_open_tickets_for_user(jira_host, user_email, headers):
    # JQL-Query, um offene Tickets eines Benutzers zu finden
    jql_query = f'assignee = "{user_email}" AND statusCategory != Done ORDER BY created DESC'
    url = f'{jira_host}/rest/api/2/search'

    # Wir fordern die gewünschten Felder explizit an
    params = {
        "jql": jql_query,
        "fields": "key,summary",  # Hier wird sowohl die Ticketnummer (key) als auch die Beschreibung (summary) angefordert
        "maxResults": 100  # Optional: Limit für die Ergebnisse
    }

    response = requests.get(url, headers=headers, params=params)

    if response.status_code == 200:
        tickets = response.json().get("issues", [])
        # Extrahiere Ticketnummer und Beschreibung
        ticket_details = [
            {"key": ticket["key"], "description": ticket["fields"].get("summary", "Keine Beschreibung")}
            for ticket in tickets
        ]
        return ticket_details
    else:
        raise Exception(f"Fehler beim Abrufen der Tickets: {response.status_code} - {response.text}")


def set_headers(sheet_name, password_reference):
    global headers

    # Hier setzen Sie die Zelle, die den session_token enthält
    jira_credentials = get_credentials(sheet_name, password_reference)
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

if __name__ == "__main__":
    fetch_jira_ticket_information()
