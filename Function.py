from openpyxl import load_workbook
from win32com import client


class FuncOfConvert:
    def __init__(self, excel_file_path, word_file_path):
        self.excel_file_path = excel_file_path
        self.word_file_path = word_file_path
        self.data = []

    def read_data_from_excel(self):
        wb = load_workbook(self.excel_file_path)
        sheet_names = wb.get_sheet_names()
        ws = wb.get_sheet_by_name(sheet_names[0])

    def win32test(self):
        excel = client.Dispatch('Excel.Application')
        word = client.Dispatch('Word.Application')

        doc = word.Documents.Open(self.word_file_path)
        book = excel.Workbooks.Open(self.excel_file_path)

        sheet = book.Worksheets(1)
        sheet.Range('A1:B8').Copy()

        wdRange = doc.Content
        wdRange.Collapse(0)

        wdRange.PasteExcelTable(False, False, False)

        doc.Save()
        doc.Close()

        book.Close()

        print('转移完成!')


if __name__ == '__main__':
    excel_path = 'C:\\Users\\Administrator\\Desktop\\123.xlsx'
    word_path = 'C:\\Users\\Administrator\\Desktop\\123.docx'

    cla = FuncOfConvert(excel_path, word_path)

    cla.win32test()
