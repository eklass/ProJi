import requests
from utils.getPasswordFrom1Password import get_credentials

from utils.Constants import JIRA_DOMAIN_CELL, VBA_SETTINGS_SHEET_NAME, CONSOLE_OUTPUT_CELL_VBA_SHEET, \
    JIRA_TICKET_DESCRIPTON_RANGE, JIRA_TICKET_NUMBER_RANGE, ONE_PASSWORD_REFERENCE_JIRA_CELL, \
    AUTOCOMPLETE_JIRA_TICKET_DESCRIPTION_CELL
from utils.excelLoader import ExcelLoader

headers = None
global_excel_loader = None

def setup_excel(sheet_name):
    excel_loader = ExcelLoader(None, None, CONSOLE_OUTPUT_CELL_VBA_SHEET)
    excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)


def fetch_jira_ticket_information():
    global headers
    setup_excel(VBA_SETTINGS_SHEET_NAME)
    get_excel_loader().log_to_excel("Started Fetching Jira Tickets for not finished Tasks assigned to this user ...", True)
    password_reference = get_excel_loader().vba_settings_sheet.range(ONE_PASSWORD_REFERENCE_JIRA_CELL).value

    # Save autocomplete flag and set it to false, to not call fetch_missing_descriptions_from_jira while executing this fetch_jira_ticket_information
    autocomplete_jira_ticket_description_before = get_excel_loader().vba_settings_sheet.range(AUTOCOMPLETE_JIRA_TICKET_DESCRIPTION_CELL).value

    get_excel_loader().vba_settings_sheet.range(AUTOCOMPLETE_JIRA_TICKET_DESCRIPTION_CELL).value = "false"

    set_headers(VBA_SETTINGS_SHEET_NAME, password_reference)
    jira_domain = get_excel_loader().vba_settings_sheet.range(JIRA_DOMAIN_CELL).value.rstrip("/")

    user_info = get_user_info(jira_domain, headers)

    tickets = get_open_tickets_for_user(user_info.get("email"), headers)
    sync_tickets_to_excel(tickets)
    get_excel_loader().log_to_excel(f"{len(tickets)} Tickets fetched from Jira")

    # restore setting
    get_excel_loader().vba_settings_sheet.range(AUTOCOMPLETE_JIRA_TICKET_DESCRIPTION_CELL).value = autocomplete_jira_ticket_description_before


def fetch_missing_descriptions_from_jira():
    global headers
    setup_excel(VBA_SETTINGS_SHEET_NAME)
    password_reference = get_excel_loader().vba_settings_sheet.range(ONE_PASSWORD_REFERENCE_JIRA_CELL).value
    set_headers(VBA_SETTINGS_SHEET_NAME, password_reference)
    jira_domain = get_excel_loader().vba_settings_sheet.range(JIRA_DOMAIN_CELL).value.rstrip("/")

    get_excel_loader().log_to_excel("Started Fetching missing jira descriptions ...", True)
    # Excel-Bereich für Ticketnummern und Beschreibungen
    description_range = get_excel_loader().vba_settings_sheet.range(JIRA_TICKET_DESCRIPTON_RANGE)
    ticket_number_range = get_excel_loader().vba_settings_sheet.range(JIRA_TICKET_NUMBER_RANGE)

    # Bestehende Daten auslesen
    existing_descriptions = [cell.value for cell in description_range]
    existing_ticket_numbers = [cell.value for cell in ticket_number_range]

    # Tickets ohne Beschreibung identifizieren
    tickets_without_description = [
        ticket_number for ticket_number, description in zip(existing_ticket_numbers, existing_descriptions)
        if ticket_number and not description
    ]

    if not tickets_without_description:
        get_excel_loader().log_to_excel("Alle Tickets haben bereits Beschreibungen.")
        return

    get_excel_loader().log_to_excel(f"Es wurden {len(tickets_without_description)} Tickets ohne Beschreibung gefunden. Abrufen der Beschreibungen...")

    # Beschreibungen über die JIRA API abrufen und ergänzen
    for ticket_number in tickets_without_description:
        url = f"{jira_domain}/rest/api/2/issue/{ticket_number}"
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            ticket_data = response.json()
            description = ticket_data.get("fields", {}).get("summary", "Keine Beschreibung")

            # Ergänze die Beschreibung in der entsprechenden Zeile
            row_index = existing_ticket_numbers.index(ticket_number) + 3  # Zeilenindex für Excel (3 = Startindex)
            get_excel_loader().vba_settings_sheet.range(f"D{row_index}").value = description
            get_excel_loader().log_to_excel(f"Beschreibung für Ticket {ticket_number} ergänzt: {description}")
        else:
            get_excel_loader().log_to_excel(f"Fehler beim Abrufen der Beschreibung für Ticket {ticket_number}: {response.status_code}")


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
        description = ticket["description"]

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


def get_open_tickets_for_user(user_email, headers):
    jira_domain = get_excel_loader().vba_settings_sheet.range(JIRA_DOMAIN_CELL).value.rstrip("/")
    # JQL-Query, um offene Tickets eines Benutzers zu finden
    jql_query = f'assignee = "{user_email}" AND statusCategory != Done ORDER BY created DESC'
    url = f'{jira_domain}/rest/api/2/search'

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
