# coding: cp1252
import traceback

from projektron import postProjektronTimesWithSelenium
from utils.Constants import COMMENT_COLUMN, DATE_CELL, HEADLESS_MODE_CELL, TIME_COLUMN, JIRA_TICKET_COLUMN, \
    PROJEKTRON_TASK_COLUMN
from utils.excelLoader import ExcelLoader, extract_time_from_cell

global_excel_loader = None


def setup_excel(sheet_name):
    excel_loader = ExcelLoader()
    excel_loader.load_excel(sheet_name)
    set_excel_loader(excel_loader)



def get_technical_task_id(non_technical_value):
    # Zugriff auf das Blatt 'VBA-Settings'
    settings_sheet = get_excel_loader().vba_settings_sheet

    # Durchlaufe die Zeilen in Spalte B ab Zeile 3 und suche nach dem nicht-technischen Wert
    for row in range(3, settings_sheet.range('B' + str(settings_sheet.cells.last_cell.row)).end('up').row + 1):
        if settings_sheet.range(f'B{row}').value == non_technical_value:
            # Wenn gefunden, gibt den korrespondierenden technischen Wert aus Spalte A zurück
            return settings_sheet.range(f'A{row}').value
    return None  # Wenn nicht gefunden, gibt None zurück

def main():
    print("Call specific method instead of main")


def post_projektron_times(sheet_name):

    try:
        excel_loader = ExcelLoader()
        time_sheet, vba_settings_sheet, wb = excel_loader.load_excel(sheet_name)
        set_excel_loader(excel_loader)
        if not time_sheet:
            get_excel_loader().log_to_excel(f"Das Blatt '{sheet_name}' wurde nicht gefunden.\n")
            return

        date = get_date(time_sheet)
        if not date:
            get_excel_loader().log_to_excel("Ungültiges Datum. Bitte überprüfen Sie das Datum.\n")
            return

        headless = get_headless_mode(vba_settings_sheet)
        tasks_to_add = collect_tasks(time_sheet)
    
        if tasks_to_add:
            try:
                response = postProjektronTimesWithSelenium.main(tasks_to_add, date, sheet_name, headless)
            except Exception as e:
                get_excel_loader().log_to_excel("\nException: " + e)
            get_excel_loader().log_to_excel("\nResponse: " + response)
            wb.save()
        else:
            get_excel_loader().log_to_excel("Kein uebereinstimmenden Projektron Task im aktuellen Sheet gefunden\n")
    except Exception as e:
        stacktrace = traceback.format_exc()  # Stacktrace als String abrufen
        get_excel_loader().log_to_excel("Fehler beim Ausführen von post_projektron_times: " + str(e) + "\n" + stacktrace)


def get_date(time_sheet):
    return time_sheet.range(DATE_CELL).value


def get_headless_mode(settings_sheet):
    return 'true' if settings_sheet.range(HEADLESS_MODE_CELL).value == 'false' else 'false'


def collect_tasks(time_sheet):
    tasks_to_add = []
    for row in range(7, 22):
        task = create_task(time_sheet, row)
        if task:
            tasks_to_add.append(task)
    return tasks_to_add


def create_task(time_sheet, row):
    duration = time_sheet.range(f'{TIME_COLUMN}{row}').value
    if duration == 0.0:
        return None

    duration_formatted = extract_time_from_cell(time_sheet.range(f'{TIME_COLUMN}{row}')) if isinstance(duration, float) else duration
    if not duration_formatted:
        return None

    ticket_number = time_sheet.range(f'{JIRA_TICKET_COLUMN}{row}').value
    description = time_sheet.range(f'{COMMENT_COLUMN}{row}').value
    human_readable_projektron_task_id = time_sheet.range(f'{PROJEKTRON_TASK_COLUMN}{row}').value
    technical_task_id = get_technical_task_id(human_readable_projektron_task_id)

    if technical_task_id:
        final_description = (ticket_number + " " if ticket_number else "") + description
        return {'task_group_oid': technical_task_id, 'duration': duration_formatted, 'description': final_description, 'row_in_timesheet': f'{row}' }
    return None


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
    post_projektron_times("Montag")
