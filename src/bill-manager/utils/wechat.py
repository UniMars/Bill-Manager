import os

# import pandas as pd
from pandas import DataFrame, read_csv, to_datetime, concat
from utils.util import BillReader, BillWriter, bill_filter, plus_minus_filter, JRZZLJ


class WCBillReader(BillReader):
    """通过继承 BillReader 类实现微信账单的读取"""

    __slots__ = "_df"

    def __init__(self, file=""):
        super().__init__(file)
        self._df = DataFrame()

    def read_file(self, filepath=""):
        """读取微信账单 CSV 文件"""
        super().read_file(filepath)
        filepath = self._file
        df = read_csv(filepath, skiprows=16)
        df['交易时间'] = to_datetime(df['交易时间'])
        df["金额(元)"] = df["金额(元)"].str.replace('¥', '').astype('float64')
        self._df = df
        return df


class WCBillWriter(BillWriter):
    __slots__ = "_ban_dict"

    def __init__(self, output_path="../template/wechat_test.xlsx", time_filter=True):
        """
        初始化 WCBillWriter 实例
        :param:
            output_path: 输出微信账单的路径，默认为当前目录下的 wechat_test.xlsx
            time_filter: 是否对时间进行过滤，即只选择符合需求的账单进行输出
        """
        super().__init__(output_path, time_filter)
        self._ban_dict = {"收/支": ['/'], "当前状态": ["已全额退款", "对方已退还", "已退款"]}

    def starter(self, path=r"D:\\Document\\材料\\账单&发票\\微信\\"):
        """
        读取指定路径下的微信账单 CSV 文件并生成 pandas.DataFrame
        :param:
            path: 微信账单 CSV 文件所在的路径，默认为 D:\\Document\\材料\\账单&发票\\微信\\
        :return:
            df_final: 合并之后的 pandas.DataFrame，包含多个微信账单的内容
        """
        df_list = []
        count = 0
        for i in super().starter(path):
            if not i.endswith(".csv"):
                continue
            count += 1
            wc = WCBillReader(file=os.path.join(path, i))
            read_df = wc.read_file()
            temp_df = DataFrame({
                '交易时间': read_df["交易时间"],
                '来源': '微信',
                '收/支': read_df['收/支'],
                '交易类型': read_df["交易类型"],
                '交易对象': read_df['交易对方'],
                '商品': read_df['商品'],
                '金额': read_df['金额(元)']})

            temp_df["母类别"] = ""
            temp_df["子类别"] = ""
            temp_df["总类别"] = ""

            temp_df["备注"] = "/当前状态：" + read_df['当前状态'] + "/交易单号：" + read_df['交易单号'] + "/商户单号：" + \
                              read_df['商户单号']
            temp_df["收支逻辑"] = read_df['收/支'].apply(plus_minus_filter)
            temp_df[JRZZLJ] = 1
            for col, ban_list in self._ban_dict.items():
                temp_df[JRZZLJ] *= (read_df[col].apply(bill_filter, args=(ban_list,)))
            df_list.append(temp_df)
        if not count:
            df_final = DataFrame(
                columns=["交易时间", "来源", "收/支", "交易类型", "交易对象", "商品", "金额", "母类别", "子类别",
                         "总类别", "备注", "收支逻辑", JRZZLJ])
            print("\n---------\n没有新的微信账单被写入")
        else:
            df_final = concat(df_list)
        df_final["乘后金额"] = df_final["收支逻辑"] * df_final[JRZZLJ] * df_final['金额']
        df_final.drop_duplicates(inplace=True)
        df_final.sort_values(by="交易时间", inplace=True)
        # self.write(df_final)
        return df_final


if __name__ == "__main__":
    a = WCBillWriter()
    a.starter()
