# -*- coding: UTF-8 -*-
from xlrd import open_workbook


class ExcelHelper:
    ADSL = {"装机地址": 0, "地址ID": 1, "外线方式": 2, "用户类型": 3 , "月租类型": 4, "账号关联账号": 5, "计费属性": 6, "密码": 7}
    ITV = {"用户类型": 8}
    WIFI = {"用户类型": 9, "场所名称": 10, "业主联系电话": 11, "安装地址": 12, "业主身份证号": 13, "安全审计合作商": 14, "客户类型": 15, "法人联系方式": 16,
            "客户经理名称": 17, "客户经理联系电话": 18, "上网方式": 19, "工商执照": 20}

    def __init__(self, path):
        self.path = path
        self.excel = open_workbook(path)
        self.sheet = self.excel.sheet_by_index(0)
        print(self.sheet)

    def get_cell_by_index(self, row, col):
        return self.sheet.cell(row, col).value

    def get_cell_value(self, type, row, col_name):
        if type != ExcelHelper.ADSL and type != ExcelHelper.WIFI and type != ExcelHelper.ITV:
            print("单元格类型不存在")
            return
        col = type[col_name]
        return self.sheet.cell(row, col).value

#excel = ExcelHelper("hotel.xls")
#print(excel.get_cell_value(excel.WIFI,2,"用户类型"))