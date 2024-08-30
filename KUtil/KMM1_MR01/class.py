class ExcelInstance():
    def __init__(self, wb=None):
        self.source_path = SOURCE_PATH
        try:
            self.app = win32com.client.gencache.EnsureDispatch('Excel.Application')
        except:
            print("Application could not be opened.")
            return
        try:
            self.open_workbook()
        except:
            print("Workbook could not be opened.")
            return
        try:
            self.ws = self.wb.Worksheets(WORKSHEET_NAME)
        except:
            print("Worksheet not found.")
            return
        self.app.Visible = True
        self.app.WindowState = win32com.client.constants.xlMaximized

    def open_workbook(self):
        """
        If it doesn't open one way, try another.
        """
        try:
            self.wb = self.app.Workbooks(self.source_path)
        except Exception as e:
            try:
                self.wb = self.app.Workbooks.Open(self.source_path)
            except Exception as e:
                print(e)
                self.wb = None

    def get_column_after(self, column, offset):
        for item in self.ws.Range("{0}{1}:{0}{2}".format(column, offset, self.get_last_row_from_column(column))).Value:
            print(item[0])

    def get_last_row_from_column(self, column):
        return self.ws.Range("{0}{1}".format(column, self.ws.Rows.Count)).End(win32com.client.constants.xlUp).Row


def main():
    f = ExcelInstance()
    f.get_column_after("A", 3)

if __name__ == "__main__":
    main()