from os.path import join

# import pandas as pd
from pandas import DataFrame, read_excel, to_datetime, concat
from xlrd import open_workbook

from utils.util import BillReader, BillWriter, bill_filter, JRZZLJ

# 设置常量
JYJE = "交易金额"
JYSJ = "交易时间"


class CCBBillReader(BillReader):
    """
    建行账单读取器
    """
    __slots__ = "_account", "_ban_dict", "_df"

    def __init__(self, file=""):
        super().__init__(file)
        # self._host_name = ""
        self._account = ""  # 银行卡号
        self._df = DataFrame()  # 账单数据

    def read_file(self, filepath=""):
        """
         读取文件
         :param filepath:
         :return:
         """
        super().read_file(filepath)
        assert self._type == "xls", "文件不是 xls 格式"
        filepath = self._file
        with open_workbook(filepath) as wb:
            sh = wb.sheet_by_index(0)
            info_row = sh.row_values(1)
            account = info_row[1][6:]  # 截取银行卡号
            # username = info_row[3][5:]
            # start_time = to_datetime(info_row[5][5:])
            # end_time = to_datetime([7][5:])
            # self._host_name = username
            self._account = account
            # title_row = sh.row_values(2)
            # print(title_row)
            df = read_excel(filepath, sheet_name=0, index_col=0, skiprows=[0, 1], header=0)
            df['交易日期'] = to_datetime(df['交易日期'], format="%Y%m%d")
            df[JYJE] = df[JYJE].str.replace(',', '').astype('float64')  # 将交易金额字符串转为浮点数
            df['账户余额'] = df['账户余额'].str.replace(',', '').astype('float64')
            self._df = df
            return df, self._account


class CCBBillWriter(BillWriter):
    """
    建行账单写入器
    """

    def __init__(self, output_path="./ccb_test.xlsx", time_filter=True):
        """
        初始化
        :param output_path: 输出文件路径
        :param time_filter: 是否启动时间过滤
        """
        super().__init__(output_path, time_filter)
        self._ban_dict = {"交易地点/附言": ["财付通", "支付宝", "微信零钱提现", "马明宇支付宝转账"],
                          '对方账号与户名': ["6217001320004689878", "6217000010127495920"],
                          "摘要": ["消费退货"]}

    def starter(self, path=r"D:\Document\材料\账单&发票\建行\\"):
        """
         开始执行
         :param path: 读取目标文件夹路径
         :return:
         """
        df_list = []
        count = 0
        for i in super().starter(path):
            if not i.endswith(".xls"):
                continue
            count += 1
            ccb = CCBBillReader(file=join(path, i))
            df, account = ccb.read_file()
            temp_df = DataFrame({
                JYSJ: df["交易日期"],
                "来源": f"建行{account}",
                "收/支": df[JYJE].apply(lambda x: "收入" if x > 0 else "支出"),
                "交易类型": df["摘要"],
                "交易对象": df['对方账号与户名'],
                "商品": "",
                "金额": df[JYJE].apply(abs),
                "母类别": "",
                "子类别": "",
                "总类别": "",
                "备注": "/" + df['交易地点/附言'],
                "收支逻辑": df[JYJE].apply(lambda x: 1 if x > 0 else -1)
            })
            temp_df[JRZZLJ] = 1
            for col in self._ban_dict:
                ban_list = self._ban_dict[col]
                temp_df[JRZZLJ] *= (df[col].apply(bill_filter, args=(ban_list,)))
            df_list.append(temp_df)
        if not count:
            df_final = DataFrame(
                columns=[JYSJ, "来源", "收/支", "交易类型", "交易对象", "商品", "金额", "母类别", "子类别",
                         "总类别",
                         "备注", "收支逻辑", JRZZLJ
                         ])  # , "乘后金额"
            print("\n---------\n没有新的建行账单被写入")
        else:
            df_final = concat(df_list)
        df_final["乘后金额"] = df_final["收支逻辑"] * df_final[JRZZLJ] * df_final['金额']
        df_final.drop_duplicates(inplace=True)
        df_final.sort_values(by=JYSJ, inplace=True)
        # self.write(df_final)
        return df_final


if __name__ == "__main__":
    a = CCBBillWriter()
    a.starter()
    print()
