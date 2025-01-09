import xlwings as xw

class ExcelLoader:
    # Konstanten
    DEFAULT_WORKBOOK_NAME = 'TimeBooking.xlsm'
    DEFAULT_VBA_SETTINGS_SHEET_NAME = 'VBA-Settings'
    DEFAULT_CONSOLE_OUTPUT_CELL = 'E27'

    def __init__(self, workbook_name=None, vba_settings_sheet_name=None, console_output_cell=None):
        self.wb = None
        self.time_sheet = None
        self.vba_settings_sheet = None
        self.workbook_name = workbook_name or ExcelLoader.DEFAULT_WORKBOOK_NAME
        self.vba_settings_sheet_name = vba_settings_sheet_name or ExcelLoader.DEFAULT_VBA_SETTINGS_SHEET_NAME
        self.console_output_cell = console_output_cell or ExcelLoader.DEFAULT_CONSOLE_OUTPUT_CELL

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

    def log_to_excel(self, log_message):
        if not self.time_sheet:
            raise ValueError("Time sheet is not loaded. Please call `load_excel` first.")

        # Aktuellen Wert der Zelle abrufen und sicherstellen, dass er ein String ist
        current_value = self.time_sheet.range(self.console_output_cell).value
        if current_value is not None:
            current_value = str(current_value)
            new_value = current_value + "\n" + log_message
        else:
            new_value = log_message

        # Setze den neuen Wert in die Zelle
        self.time_sheet.range(self.console_output_cell).value = new_value

        # Konsolenausgabe für Debugging
        print(log_message)
