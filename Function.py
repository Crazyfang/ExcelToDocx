from openpyxl import load_workbook
import re


class FuncOfConvert:
    def __init__(self, excel_file_path, word_file_path):
        self.excel_file_path = excel_file_path
        self.word_file_path = word_file_path
        self.data = {}

    def read_data_from_excel(self):
        wb = load_workbook(self.excel_file_path)
        sheet_names = wb.sheetnames
        ws = wb[sheet_names[0]]

        # 匹配支行和客户经理
        partten_manager_name = re.compile(r'(?P<name>.*)：(?P<bank_name>.*)（(?P<manager_person>.*)）')
        match = partten_manager_name.match(ws.cell(row=3, column=1).value.strip())

        self.data['bank_name'] = match.group('bank_name')  # 支行名称 A3's value
        self.data['customer_name'] = ws.cell(row=4, column=2).value  # 客户 B4's value
        self.data['manager_person'] = match.group('manager_person')  # 归管客户经理 A3's value
        self.data['customer_basic_info_1'] = ws.cell(row=31, column=1).value  # 借款人基本情况1
        self.data['customer_basic_info_2'] = ws.cell(row=32, column=1).value  # 借款人基本情况2
        self.data['associate_enterprise_info'] = ws.cell(row=34, column=1).value  # 关联企业情况
        self.data['associate_merge_table'] = ws.cell(row=35, column=1).value  # 关联并表
        self.data['enterprise_operator_info_1'] = '\n'.join(
            ws.cell(row=49, column=1).value.split('\n')[1:])  # 企业经营者相关情况1
        self.data['enterprise_operator_info_2'] = '\n'.join(
            ws.cell(row=50, column=1).value.split('\n')[1:])  # 企业经营者相关情况2
        self.data['enterprise_finance_condition_1'] = '\n'.join(
            ws.cell(row=33, column=1).value.split('\n')[1:])  # 企业财务状况1
        self.data['enterprise_finance_condition_2'] = ws.cell(row=158, column=1).value  # 企业财务状况2
        self.data['enterprise_finance_condition_3'] = ws.cell(row=159, column=1).value  # 企业财务状况3
        self.data['warrantor_and_guaranty_1'] = '\n'.join(ws.cell(row=87, column=1).value.split('\n')[1:])  # 保证人及抵押物情况
        self.data['warrantor_and_guaranty_2'] = '\n'.join(ws.cell(row=88, column=1).value.split('\n')[1:])  # 保证人及抵押物情况
        self.data['declaration_reason_and_purpose_1'] = '\n'.join(
            ws.cell(row=168, column=1).value.split('\n')[1:])  # 支行申报理由及用途1
        self.data['declaration_reason_and_purpose_2'] = '\n'.join(
            ws.cell(row=169, column=1).value.split('\n')[1:])  # 支行申报理由及用途2

    def win32test(self):
        excel = client.Dispatch('Excel.Application')
        word = client.Dispatch('Word.Application')

        doc = word.Documents.Open(self.word_file_path)
        book = excel.Workbooks.Open(self.excel_file_path)

        sheet = book.Worksheets(1)
        sheet.Range('A36:AE44').Copy()

        myRange = doc.Range()
        myRange.InsertAfter(self.data['bank_name'] + '：' + self.data['customer_name'] + '  归管客户经理：' + self.data['manager_person'])
        myRange.InsertAfter('(一)借款人基本情况')
        myRange.InsertAfter(self.data['customer_basic_info_1'])
        myRange.InsertAfter(self.data['customer_basic_info_2'])
        myRange.InsertAfter('关联企业情况：' + self.data['associate_enterprise_info'])
        myRange.InsertAfter('关联并表：' + self.data['associate_merge_table'])
        myRange.InsertAfter('(二)企业经营者相关情况')
        myRange.InsertAfter(self.data['enterprise_operator_info_1'])
        myRange.InsertAfter(self.data['enterprise_operator_info_2'])

        wdRange = doc.Content
        wdRange.Collapse(0)
        wdRange.PasteExcelTable(False, False, False)
        wdRange.Collapse(0)

        myRange.InsertAfter('(三)企业财务状况')
        myRange.InsertAfter(self.data['enterprise_finance_condition_1'])

        sheet.Range('A108:AE157').Copy()
        wdRange.PasteExcelTable(False, False, False)
        wdRange.Collapse(0)

        myRange.InsertAfter(self.data['enterprise_finance_condition_2'])
        myRange.InsertAfter(self.data['enterprise_finance_condition_3'])

        myRange.InsertAfter('(四)存量授信及申报授信情况')

        sheet.Range('A62:AE86').Copy()
        wdRange.PasteExcelTable(False, False, False)
        wdRange.Collapse(0)

        myRange.InsertAfter('保证人及抵押物情况介绍')
        myRange.InsertAfter(self.data['warrantor_and_guaranty_1'])
        myRange.InsertAfter(self.data['warrantor_and_guaranty_2'])

        myRange.InsertAfter('第二保证人落实情况：')
        sheet.Range('A161:AE166').Copy()
        wdRange.PasteExcelTable(False, False, False)
        wdRange.Collapse(0)

        myRange.InsertAfter('(五)支行申报理由及用途')
        myRange.InsertAfter(self.data['declaration_reason_and_purpose_1'])
        myRange.InsertAfter(self.data['declaration_reason_and_purpose_2'])

        myRange.InsertAfter('授信部意见：')
        myRange.InsertAfter('风险提示：')
        myRange.InsertAfter('(六)授信审批委员会集体审议结论')


        myRange.InsertAfter(self.data['customer_name'])

        doc.Save()
        doc.Close()

        book.Close()

        print('转移完成!')

    def test(self):
        # partten = re.compile(r'(?P<name>.*)：(?P<bank_name>.*)（(?P<manager_person>.*)）')
        # string = '申报单位：招宝山支行（方勇）              '
        partten = re.compile(r'(?P<name>.*)\s(?P<bank_name>.*)')
        string = '法定代表人（实际经营者）及配偶相关说明：（个人情况、家庭成员）。\n  法定代表人****** \n 1233213'
        print(string)
        match = partten.match(string)
        print(match.group('name'))
        print(match.group('bank_name'))
        # print(match.group('manager_person'))
        print(match[0])


if __name__ == '__main__':
    excel_path = 'C:\\Users\\Administrator\\Desktop\\123.xlsx'
    word_path = 'C:\\Users\\Administrator\\Desktop\\123.docx'

    cla = FuncOfConvert(excel_path, word_path)
    cla.test()
    # cla.win32test()
