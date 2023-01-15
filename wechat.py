import pandas as pd

from utils import BillReader, BillWriter, bill_filter

JRZZ = "计入总账逻辑"


def plus_minus_filter(x):
    if x == "收入":
        return 1
    elif x == "支出":
        return -1
    else:
        return 0


class WCBillReader(BillReader):
    __slots__ = "_df"

    def __init__(self, file=""):
        super().__init__(file)
        self._df: pd.DataFrame

    def read_file(self, filepath=""):
        super().read_file(filepath)
        filepath = self._file
        df = pd.read_csv(filepath, skiprows=16)
        df['交易时间'] = pd.to_datetime(df['交易时间'])
        df["金额(元)"] = df["金额(元)"].str.replace('¥', '').astype('float64')
            # .map(lambda x: '%.2f' % x)

        # , index_col = 0, skiprows = [0, 15], header = 0, error_bad_lines=False
        return df


class WCBillWriter(BillWriter):
    __slots__ = "_ban_dict"

    def __init__(self, output_path="./wechat_test.xlsx", time_filter=True):
        super().__init__(output_path, time_filter)
        # self._output_path = "./wechat_test.xlsx"
        self._ban_dict = {"收/支": ['/'],
                          "当前状态": ["已全额退款", "对方已退还", "已退款"]}

    def starter(self, path=r"D:\\Document\\材料\\账单&发票\\微信\\"):
        df_list = []
        count = 0
        for i in super().starter(path):
            suffix = (path + i).split('.')[-1]
            if suffix != "csv":
                continue
            count += 1
            wc = WCBillReader(file=path + i)
            df = wc.read_file()
            df1 = pd.DataFrame()
            df1["交易时间"] = df["交易时间"]
            df1["来源"] = "微信"
            df1["收/支"] = df['收/支']
            df1["交易类型"] = df["交易类型"]
            df1["交易对象"] = df['交易对方']
            df1["商品"] = df['商品']
            df1["金额"] = df['金额(元)']

            # TODO 类别分类
            df1["母类别"] = ""
            df1["子类别"] = ""
            df1["总类别"] = ""

            df1["备注"] = df['备注'] + "/交易单号：" + df['交易单号'] + "/当前状态：" + df['当前状态']
            df1["收支逻辑"] = df['收/支'].apply(plus_minus_filter)
            df1[JRZZ] = 1
            for col in self._ban_dict:
                ban_list = self._ban_dict[col]
                df1[JRZZ] *= (df[col].apply(bill_filter, args=(ban_list,)))
            df1["乘后金额"] = df1['金额'] * df1[JRZZ] * df1['收支逻辑']
            df_list.append(df1)
        if not count:
            df_final = pd.DataFrame(
                columns=["交易时间", "来源", "收/支", "交易类型", "交易对象", "商品", "金额", "母类别", "子类别",
                         "总类别",
                         "备注", "收支逻辑", "计入总账逻辑", "乘后金额"
                         ])
            print("\n---------\nno new bill written")
        else:
            df_final = pd.concat(df_list)
        # df11 = df_final[df_final[JRZZ] == 1]
        df_final.drop_duplicates(inplace=True)
        df_final = df_final.sort_values(by="交易时间")
        # self.write(df_final)
        return df_final

    def write(self, df: pd.DataFrame):
        book, df, sh = super().write(df)
        df = df[df[JRZZ] == 1]
        for index, row in df.iterrows():
            row_list = list(row)
            sh.append(row_list)
            # print()
        book.save(self._output_path)
        print("\n写入完成")


if __name__ == "__main__":
    # a = WCBillReader(r"D:\Document\材料\账单&发票\微信\微信支付账单(20220113-20220413).csv")
    # #     # a.read_file()

    a = WCBillWriter()
    a.starter()
