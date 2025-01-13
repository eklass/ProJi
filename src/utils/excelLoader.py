import xlwings as xw

from utils.Constants import WORKBOOK_NAME, VBA_SETTINGS_SHEET_NAME, CONSOLE_OUTPUT_CELL, CONSOLE_OUTPUT_CELL_VBA_SHEET


class ExcelLoader:
    # Konstanten

    def __init__(self, workbook_name=None, vba_settings_sheet_name=None, console_output_cell=None):
        self.wb = None
        self.time_sheet = None
        self.vba_settings_sheet = None
        self.workbook_name = workbook_name or WORKBOOK_NAME
        self.vba_settings_sheet_name = vba_settings_sheet_name or VBA_SETTINGS_SHEET_NAME
        self.console_output_cell = console_output_cell or CONSOLE_OUTPUT_CELL

    def load_excel(self, sheet_name):
        if not self.wb:
            try:
                self.wb = xw.Book.caller()
            except Exception:
                # An exception would occur, if we try to debug the code via IDE and therefore not call the Script via Excel
                self.wb = xw.Book(self.workbook_name)

        if not self.time_sheet or not self.vba_settings_sheet:
            self.time_sheet = self.wb.sheets[sheet_name]
            self.vba_settings_sheet = self.wb.sheets[self.vba_settings_sheet_name]
        return self.time_sheet, self.vba_settings_sheet, self.wb

    def get_sheet(self, sheet_name):
        if self.wb:
            return self.wb.sheets[sheet_name]
        else:
            raise ValueError("Workbook is not loaded. Please call `load_excel` first.")

    def get_time_sheet(self):
        if self.time_sheet:
            return self.time_sheet
        else:
            raise ValueError("Workbook is not loaded. Please call `load_excel` first.")

    def log_to_excel(self, log_message, clear=False):
        if not self.time_sheet:
            raise ValueError("Time sheet is not loaded. Please call `load_excel` first.")

        # Aktuellen Wert der Zelle abrufen und sicherstellen, dass er ein String ist
        old_console_output_cell = self.console_output_cell
        if (self.time_sheet.name == VBA_SETTINGS_SHEET_NAME):
            self.console_output_cell = CONSOLE_OUTPUT_CELL_VBA_SHEET
        current_value = self.time_sheet.range(self.console_output_cell).value
        if current_value is not None and clear == False:
            current_value = str(current_value)
            new_value = current_value + "\n" + log_message
        else:
            new_value = log_message

        # Setze den neuen Wert in die Zelle
        self.time_sheet.range(self.console_output_cell).value = new_value

        self.console_output_cell = old_console_output_cell
        # Konsolenausgabe für Debugging
        print(log_message)


def extract_time_from_cell(cell):
    cell_value = cell.value
    try:
        # Versuche, den Wert in einen float zu konvertieren
        hours = float(cell_value) * 24  # Zelle wird in Stunden umgerechnet
        return hours  # Gib die Stunden als Float zurück
    except (ValueError, TypeError):
        # Gib eine sinnvolle Fehlermeldung oder einen Fallback zurück, falls die Konvertierung fehlschlägt
        return None


def convert_time_to_decimal(time_str):
    """
    Konvertiert eine Zeit im Format hh:mm in Dezimalstunden.

    :param time_str: Zeit im Format hh:mm (z.B. '00:15')
    :return: Dezimaldarstellung der Stunden (z.B. 0.25)
    """
    if not time_str or ":" not in time_str:
        return None

    hours, minutes = map(int, time_str.split(":"))
    return hours + minutes / 60.0
