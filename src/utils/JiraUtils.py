from utils.Constants import PROJI_SETTINGS_JIRA_TICKET_MAPPING_COLUMN, PROJI_SETTINGS_ONE_PASSWORD_REFERENCE_JIRA_COLUMN, PROJI_SETTINGS_JIRA_HOST_COLUMN
from utils.excelLoader import ExcelLoader
from utils.getPasswordFrom1Password import get_credentials
import base64

global_excel_loader = None

def get_jira_mapping_for_ticket_number(ticket_number, proJiSettingsSheet):
    # Iteriere über die Zeilen 7 bis 16
    for row in range(7, 17):
        ticketPrefix = proJiSettingsSheet.range(f'{PROJI_SETTINGS_JIRA_TICKET_MAPPING_COLUMN}{row}').value
        if ticketPrefix is None:
            raise TypeError("Keine Einstellung in Proji-Settings für " + ticket_number +" gefunden")
        ticketPrefixes = ticketPrefix.split(',')  # Trenne den string in eine Liste

        # Überprüfe, ob eines der Präfixe in der ticket_number enthalten ist
        for prefix in ticketPrefixes:
            if prefix.strip() in ticket_number:  # .strip() entfernt mögliche Leerzeichen
                # Falls ja, gib den Wert aus der entsprechenden Spalte zurück
                return row

    return None  # Oder ein spezieller Wert oder eine Exception, je nach Anwendungsfall


def get_jira_host_with_row(proJiSettingsSheet):
    # Iteriere über die Zeilen 7 bis 16
    jira_hosts_with_rows = []
    for row in range(7, 17):
        jira_host = proJiSettingsSheet.range(f'{PROJI_SETTINGS_JIRA_HOST_COLUMN}{row}').value
        if jira_host is not None:
            jira_hosts_with_rows.append((jira_host, row))

    return jira_hosts_with_rows


def get_password_reference(current_jira_mapping_row):
    proji_settings_sheet = get_excel_loader().get_sheet("ProJi-Settings")
    password_reference = proji_settings_sheet.range(
        f'{PROJI_SETTINGS_ONE_PASSWORD_REFERENCE_JIRA_COLUMN}{current_jira_mapping_row}').value
    return password_reference

def get_jira_domain(current_jira_mapping_row):
    proji_settings_sheet = get_excel_loader().get_sheet("ProJi-Settings")
    jira_domain = proji_settings_sheet.range(f'{PROJI_SETTINGS_JIRA_HOST_COLUMN}{current_jira_mapping_row}').value.rstrip("/")
    return jira_domain


def get_worklog_url_for_ticket_number(ticket_number, current_jira_mapping_row):
    jira_domain = get_jira_domain(current_jira_mapping_row)
    url_template = f'{jira_domain}/rest/api/2/issue/{{issue_key}}/worklog'
    worklog_url = url_template.format(issue_key=ticket_number.strip())
    return worklog_url


#TODO: Optimize Getter and Setter to bundle those calls in one class
def set_excel_loader(excel_loader):
    global global_excel_loader
    if not isinstance(excel_loader, ExcelLoader):
        raise TypeError("Das übergebene Objekt ist kein ExcelLoader.")
    global_excel_loader = excel_loader


def get_excel_loader():
    global global_excel_loader
    if global_excel_loader is None:
        excel_loader = ExcelLoader()
        excel_loader.load_excel("ProJi-Settings")
        set_excel_loader(excel_loader)
    return global_excel_loader


def create_http_header(sheet_name, password_reference):
    # Hier setzen Sie die Zelle, die den session_token enthält
    jira_credentials = get_credentials(sheet_name, password_reference)
    personal_access_token = jira_credentials['password']
    username = jira_credentials['username']
    if not personal_access_token:
        get_excel_loader().log_to_excel("Error during fetching personal accessToken from 1Password in checkJiraTimes")
        return

    #     # Basic Auth: Combine "username:api_token" and encode it in Base64
    auth_string = f"{username}:{personal_access_token}"
    encoded_auth = base64.b64encode(auth_string.encode('utf-8')).decode('utf-8')
    headers = {'Authorization': f'Basic {encoded_auth}'}
    return headers
