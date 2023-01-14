import re
import pandas as pd
import xlrd
from utils import BillReader, BillWriter

JYJE = "交易金额"

JRZZ = "计入总账逻辑"


def start_filter(word, ban_list: list):
    if not isinstance(word, str):
        return True
    if re.split('[-/]', word)[0] in ban_list:
        return False
    return True


class CCBBillReader(BillReader):
    __slots__ = "_account", "_ban_dict", "_df"

    def __init__(self, file=""):
        super().__init__(file)
        self._host_name = ""
        self._account = ""
        self._df = pd.DataFrame()
        # if file != "":
        #     self.read_file()

    def read_file(self, filepath=""):
        super().read_file(filepath)
        assert self._type == "xls", "file is not xls"
        filepath = self._file
        with xlrd.open_workbook(filepath) as wb:
            # nsheets = wb.sheet_names()
            sh = wb.sheet_by_index(0)
            info_row = sh.row_values(1)
            account = info_row[1][6:]
            username = info_row[3][5:]
            # starttime = pd.to_datetime(info_row[5][5:])
            # endtime = pd.to_datetime([7][5:])
            self._host_name = username
            self._account = account
            title_row = sh.row_values(2)
            print(title_row)
            df = pd.read_excel(filepath, sheet_name=0, index_col=0, skiprows=[0, 1], header=0)
            df['交易日期'] = pd.to_datetime(df['交易日期'], format="%Y%m%d")
            df[JYJE] = df[JYJE].str.replace(',', '').astype('float64').map(lambda x: '%.2f' % x)
            df['账户余额'] = df['账户余额'].str.replace(',', '').astype('float64').map(lambda x: '%.2f' % x)
            self._df = df
            return df, self._account


class CCBBillWriter(BillWriter):
    def __init__(self, output_path=r"C:\Users\Lenovo\OneDrive\Documents\个人财务管理.xlsx", time_filter=True):
        super().__init__(output_path, time_filter)
        self._ban_dict = {"交易地点/附言": ["财付通", "支付宝"],
                          '对方账号与户名': ["6217001320004689878", "6217000010127495920"]}

    def starter(self, path=r"D:\Document\材料\账单&发票\建行\\"):
        df_list = []
        for i in super().starter(path):
            suffix = (path + i).split('.')[-1]
            if suffix != "csv" and suffix != "xlsx" and suffix != "xls":
                continue
            ccb = CCBBillReader(file=path + i)
            df, account = ccb.read_file()
            df1 = pd.DataFrame()
            df1["交易时间"] = df["交易日期"]
            df1["来源"] = "建行" + account
            df1["收/支"] = df[JYJE].apply(lambda x: "收入" if float(x) > 0 else "支出")
            # df1["交易类型"]=1
            df1["交易类型"] = df["摘要"]
            df1["交易对象"] = df['对方账号与户名']
            df1["商品"] = ""
            df1["金额"] = df[JYJE].apply(float).apply(abs)
            df1["母类别"] = ""
            df1["子类别"] = ""
            df1["总类别"] = ""
            df1["备注"] = df['交易地点/附言']
            df1["收支逻辑"] = df[JYJE].apply(lambda x: 1 if float(x) > 0 else -1)
            df1[JRZZ] = 1
            for col in self._ban_dict:
                ban_list = self._ban_dict[col]
                df1[JRZZ] *= (df[col].apply(start_filter, args=(ban_list,)))
            df1["乘后金额"] = df1['金额'] * df1[JRZZ]
            df_list.append(df1)
        df_final = pd.concat(df_list)
        # self.write(df_final)
        return df_final

    # def write(self, df: pd.DataFrame):
    #     book, df, sh = super().write(df)
    #     df11 = df[df[JRZZ] == 1]
    #     for index, row in df11.iterrows():
    #         row_list = list(row)
    #         sh.append(row_list)
    #         # print()
    #     book.save(self._output_path)
    #     print("写入完成")


if __name__ == "__main__":
    a = CCBBillWriter()
    a.starter()
    print()
