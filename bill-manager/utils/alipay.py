# coding=utf-8
import os
from parser import ParserError
from re import compile
from openpyxl import load_workbook
import pandas as pd
from utils.util import BillReader, BillWriter, plus_minus_filter, JRZZLJ, bill_filter


class ALiPBillReader(BillReader):
    __slots__ = "_df"

    def __init__(self, file=""):
        super().__init__(file)
        self._df = pd.DataFrame()

    def read_file(self, filepath=""):
        super().read_file(filepath)
        filepath = self._file
        with open(filepath, 'r', encoding='gb2312') as f:
            lines = f.readlines()
            if lines[0].startswith("------------------------支付宝（中国）网络技术有限公司  电子客户回单"):
                try:
                    df = pd.read_csv(filepath, skiprows=1, skipfooter=21, encoding="gb2312", engine="python")
                except Exception as e:
                    print(e)
            elif lines[0].startswith(
                    "------------------------------------------------------------------------------------"):
                try:
                    df = pd.read_csv(filepath, skiprows=24, encoding="gb2312")
                except Exception as e:
                    print(e)
            else:
                raise ParserError("支付宝账单读取失败")

        # new_columns = list(map(lambda t: t.strip(), df.columns))
        new_columns = df.columns.str.strip()
        df.columns = new_columns
        df['交易时间'] = pd.to_datetime(df['交易时间'])
        df["金额"] = df["金额"].astype('float64')
        self._df = df
        return df


class ALiPBillWriter(BillWriter):
    __slots__ = "_ban_dict", "_refund_df"

    def __init__(self, output_path=r"./alipay_test.xlsx", time_filter=True):
        super().__init__(output_path, time_filter)
        self._refund_df = pd.DataFrame()
        self._ban_dict = {"交易状态": ["退款成功", "交易关闭", "还款成功"]}

    # 功能：读取和处理支付宝账单数据
    def starter(self, path=r"D:\\Document\\材料\\账单&发票\\支付宝\\"):
        """
        读取和处理支付宝账单数据
        :param path: 支付宝账单所在的根文件夹路径
        :return: 经过处理的 pandas DataFrame，包含如下列："交易时间"，"来源"，"收/支"，"交易类型"，"交易对象"，"商品"，"金额"，"母类别"，"子类别"，"总类别"，"备注"，"收支逻辑"，"计入总账逻辑"。
        """
        df_list = []
        refund_df = pd.DataFrame()
        count = 0
        for i in super().starter(path):
            if not i.endswith(".csv"):
                continue
            count += 1
            ali = ALiPBillReader(file=os.path.join(path, i))
            df = ali.read_file()
            temp_df = pd.DataFrame()
            temp_df["交易时间"] = df["交易时间"]
            temp_df["来源"] = "支付宝"
            temp_df["收/支"] = df['收/支']
            temp_df["交易类型"] = df["交易分类"]
            temp_df["交易对象"] = df['交易对方'] + '/' + df['对方账号']
            temp_df["商品"] = df['商品说明']
            temp_df["金额"] = df['金额']

            temp_df["母类别"] = ""
            temp_df["子类别"] = ""
            temp_df["总类别"] = ""
            temp_df["备注"] = "/交易状态：" + df['交易状态'] + "/交易订单号：" + df['交易订单号'] + "/商家订单号：" + df[
                '商家订单号']

            refund_df = pd.concat([refund_df, df[df["交易状态"].str.startswith("退款成功")]])

            temp_df["收支逻辑"] = df['收/支'].apply(plus_minus_filter)
            temp_df[JRZZLJ] = 1
            for col, ban_list in self._ban_dict.items():
                temp_df[JRZZLJ] *= (df[col].apply(bill_filter, args=(ban_list,)))
            df_list.append(temp_df)
        if len(refund_df):
            refund_df["original_交易订单号"] = refund_df["交易订单号"].str.split(compile("[_*]")).apply(lambda x: x[0])
        if not count:
            df_final = pd.DataFrame(
                columns=["交易时间", "来源", "收/支", "交易类型", "交易对象", "商品", "金额", "母类别", "子类别",
                         "总类别",
                         "备注", "收支逻辑", JRZZLJ
                         ])  # , "乘后金额"
            print("\n---------\n没有新的支付宝账单被写入")
        else:
            df_final = pd.concat(df_list)
        df_final["乘后金额"] = df_final["收支逻辑"] * df_final[JRZZLJ] * df_final['金额']
        df_final.drop_duplicates(inplace=True)
        df_final.sort_values(by="交易时间", inplace=True)
        # self.write(df_final)
        self._refund_df = refund_df
        return df_final

    # 功能：将退款数据写入支付宝账单
    def refund_write(self):
        """
        将退款数据写入支付宝账单
        """
        if not len(self._refund_df):
            print("\n没有退款\n")
            return
        book = load_workbook(self._output_path)
        sh = book['流水表']
        column_source = tuple(map(lambda c: c.value, sh["B"]))
        column_memo = sh["K"]
        cell_df = pd.DataFrame([column_source, column_memo]).T
        cell_df = cell_df[cell_df[0].notnull()]
        memo = cell_df[cell_df[0].str.startswith("支付宝")].iloc[::-1][1]
        for _, row in self._refund_df.iterrows():
            po_no = row["original_交易订单号"]
            amount = row['金额']
            final_row_no = -1
            pattern = compile(f"交易订单号：{po_no}")
            for cell in memo:
                try:
                    if pattern.search(cell.value):
                        final_row_no = cell.row
                        break
                except Exception as e:
                    print(e)
            if final_row_no != -1:
                final_row = sh[final_row_no]
                amount_cell = final_row[6]
                amount_cell.value -= amount
                if amount_cell.value == 0:
                    final_row[12].value = 0
        book.save(self._output_path)
        print("\n退款写入完成")


if __name__ == "__main__":
    a = ALiPBillWriter()
    a.starter()
