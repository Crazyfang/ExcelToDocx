from win32com import client
from openpyxl import load_workbook
import re
import os
import xlrd


class FuncOfConvert:
    def __init__(self):
        self.data = {}
        self.excel = client.Dispatch('Excel.Application')
        self.word = client.Dispatch('Word.Application')

    def get_file_list(self, file_path):
        file_list = []
        for file_name in os.listdir(file_path):
            extension = os.path.splitext(file_name)[-1][1:]
            if extension == 'xlsx' or extension == 'xls':
                file_list.append(os.path.join(file_path, file_name))

        return file_list

    def read_data_from_excel(self, excel_file_path):
        wb = load_workbook(excel_file_path, read_only=True)
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

    def read_data_from_xls(self, excel_file_path):
        excel = xlrd.open_workbook(excel_file_path)
        ws = excel.sheet_by_index(0)

        # 匹配支行和客户经理
        partten_manager_name = re.compile(r'(?P<name>.*)：(?P<bank_name>.*)（(?P<manager_person>.*)）')
        match = partten_manager_name.match(ws.cell(rowx=2, colx=0).value.strip())

        self.data['bank_name'] = match.group('bank_name')  # 支行名称 A3's value
        self.data['customer_name'] = ws.cell(rowx=3, colx=1).value  # 客户 B4's value
        self.data['manager_person'] = match.group('manager_person')  # 归管客户经理 A3's value
        self.data['customer_basic_info_1'] = ws.cell(rowx=30, colx=0).value  # 借款人基本情况1
        self.data['customer_basic_info_2'] = ws.cell(rowx=31, colx=0).value  # 借款人基本情况2
        self.data['associate_enterprise_info'] = ws.cell(rowx=33, colx=0).value  # 关联企业情况
        self.data['associate_merge_table'] = ws.cell(rowx=34, colx=0).value  # 关联并表
        self.data['enterprise_operator_info_1'] = '\n'.join(
            ws.cell(rowx=48, colx=0).value.split('\n')[1:])  # 企业经营者相关情况1
        self.data['enterprise_operator_info_2'] = '\n'.join(
            ws.cell(rowx=49, colx=0).value.split('\n')[1:])  # 企业经营者相关情况2
        self.data['enterprise_finance_condition_1'] = '\n'.join(
            ws.cell(rowx=32, colx=0).value.split('\n')[1:])  # 企业财务状况1
        self.data['enterprise_finance_condition_2'] = ws.cell(rowx=157, colx=0).value  # 企业财务状况2
        self.data['enterprise_finance_condition_3'] = ws.cell(rowx=158, colx=0).value  # 企业财务状况3
        self.data['warrantor_and_guaranty_1'] = '\n'.join(ws.cell(rowx=86, colx=0).value.split('\n')[1:])  # 保证人及抵押物情况
        self.data['warrantor_and_guaranty_2'] = '\n'.join(ws.cell(rowx=87, colx=0).value.split('\n')[1:])  # 保证人及抵押物情况
        self.data['declaration_reason_and_purpose_1'] = '\n'.join(
            ws.cell(rowx=167, colx=0).value.split('\n')[1:])  # 支行申报理由及用途1
        self.data['declaration_reason_and_purpose_2'] = '\n'.join(
            ws.cell(rowx=168, colx=0).value.split('\n')[1:])  # 支行申报理由及用途2

    def win32test(self, excel_file_path):
        doc_file_path = os.path.splitext(excel_file_path)[0] + '.docx'
        doc = self.word.Documents.Add()
        book = self.excel.Workbooks.Open(excel_file_path)

        sheet = book.Worksheets(1)
        sheet.Range('A36:AE44').Copy()

        # myRange = doc.Range()
        # myRange = doc.Selection
        self.word.Selection.InsertAfter(self.data['bank_name'] + '：' + self.data['customer_name'] + '  归管客户经理：' + self.data[
            'manager_person'] + '\n')
        self.word.Selection.InsertAfter('(一)借款人基本情况\n')
        self.word.Selection.InsertAfter(self.data['customer_basic_info_1'])
        self.word.Selection.InsertAfter(self.data['customer_basic_info_2'])
        self.word.Selection.InsertAfter('关联企业情况：' + self.data['associate_enterprise_info'])
        self.word.Selection.InsertAfter('关联并表：' + self.data['associate_merge_table'])
        self.word.Selection.InsertAfter('(二)企业经营者相关情况\n')
        self.word.Selection.InsertAfter(self.data['enterprise_operator_info_1'])
        self.word.Selection.InsertAfter(self.data['enterprise_operator_info_2'])

        # wdRange = doc.Content
        # wdRange.Collapse(0)
        self.word.Selection.MoveRight()
        self.word.Selection.PasteExcelTable(False, False, False)
        # wdRange.PasteExcelTable(False, False, False)

        self.word.Selection.InsertAfter('(三)企业财务状况\n')
        self.word.Selection.InsertAfter(self.data['enterprise_finance_condition_1'])

        sheet.Range('A108:AE157').Copy()
        self.word.Selection.MoveRight()
        self.word.Selection.PasteExcelTable(False, False, False)
        # wdRange.PasteExcelTable(False, False, False)

        self.word.Selection.InsertAfter(self.data['enterprise_finance_condition_2'])
        self.word.Selection.InsertAfter(self.data['enterprise_finance_condition_3'])

        self.word.Selection.InsertAfter('(四)存量授信及申报授信情况\n')

        sheet.Range('A62:AE86').Copy()
        self.word.Selection.MoveRight()
        self.word.Selection.PasteExcelTable(False, False, False)
        # wdRange.PasteExcelTable(False, False, False)

        self.word.Selection.InsertAfter('保证人及抵押物情况介绍\n')
        self.word.Selection.InsertAfter(self.data['warrantor_and_guaranty_1'])
        self.word.Selection.InsertAfter(self.data['warrantor_and_guaranty_2'])

        self.word.Selection.InsertAfter('第二保证人落实情况：\n')
        sheet.Range('A161:AE166').Copy()
        self.word.Selection.MoveRight()
        self.word.Selection.PasteExcelTable(False, False, False)
        # wdRange.PasteExcelTable(False, False, False)

        self.word.Selection.InsertAfter('(五)支行申报理由及用途\n')
        self.word.Selection.InsertAfter(self.data['declaration_reason_and_purpose_1'])
        self.word.Selection.InsertAfter(self.data['declaration_reason_and_purpose_2'])

        self.word.Selection.InsertAfter('授信部意见：\n')
        self.word.Selection.InsertAfter('风险提示：\n')
        self.word.Selection.InsertAfter('(六)授信审批委员会集体审议结论\n')

        doc.SaveAs(doc_file_path)
        doc.Close()

        book.Application.CutCopyMode = False
        book.Close()

        self.data.clear()

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
    excel_path = 'C:\\Users\\Administrator\\Desktop\\2020审核调查报告—对公.xlsx'
    word_path = 'C:\\Users\\Administrator\\Desktop\\123.docx'

    cla = FuncOfConvert(excel_path, word_path)
    cla.read_data_from_excel()
    cla.win32test()
    # cla.test()
    # cla.win32test()
