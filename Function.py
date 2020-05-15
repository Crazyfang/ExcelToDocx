from openpyxl import load_workbook


class FuncOfConvert:
    def __init__(self, excel_file_path, word_file_path):
        self.excel_file_path = excel_file_path
        self.word_file_path = word_file_path
        self.data = {}

    def read_data_from_excel(self):
        wb = load_workbook(self.excel_file_path)
        sheet_names = wb.get_sheet_names()
        ws = wb.get_sheet_by_name(sheet_names[0])

        self.data['bank_name'] = ws.cell(row=3, column=1).value           # 支行名称 A3's value
        self.data['customer_name'] = ws.cell(row=4, column=2).value       # 客户 B4's value
        self.data['manager_person'] = ws.cell(row=3, column=1).value      # 归管客户经理 A3's value
        self.data['customer_basic_info_1'] = ws.cell(row=31, column=1).value    # 借款人基本情况1
        self.data['customer_basic_info_2'] = ws.cell(row=32, column=1).value    # 借款人基本情况2
        self.data['associate_enterprise_info'] = ws.cell(row=34, column=1).value    # 关联企业情况
        self.data['associate_merge_table'] = ws.cell(row=35, column=1).value        # 关联并表
        self.data['enterprise_operator_info_1'] = ws.cell(row=49, column=1).value     # 企业经营者相关情况1
        self.data['enterprise_operator_info_2'] = ws.cell(row=50, column=1).value     # 企业经营者相关情况2
        self.data['enterprise_finance_condition_1'] = ws.cell(row=33, column=1).value   # 企业财务状况1
        self.data['enterprise_finance_condition_2'] = ws.cell(row=158, column=1).value  # 企业财务状况2
        self.data['enterprise_finance_condition_3'] = ws.cell(row=159, column=1).value  # 企业财务状况3
        self.data['warrantor_and_guaranty_1'] = ws.cell(row=87, column=1).value           # 保证人及抵押物情况
        self.data['warrantor_and_guaranty_2'] = ws.cell(row=88, column=1).value         # 保证人及抵押物情况
        self.data['declaration_reason_and_purpose_1'] = ws.cell(row=168, column=1).value # 支行申报理由及用途1
        self.data['declaration_reason_and_purpose_2'] = ws.cell(row=169, column=1).value  # 支行申报理由及用途2


if __name__ == '__main__':
    excel_path = 'C:\\Users\\Administrator\\Desktop\\123.xlsx'
    word_path = 'C:\\Users\\Administrator\\Desktop\\123.docx'

    cla = FuncOfConvert(excel_path, word_path)

    # cla.win32test()
