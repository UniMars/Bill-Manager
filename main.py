"""
Author       : UniMars
Date         : 2023-01-13 23:19:59
LastEditors  : UniMars
LastEditTime : 2023-03-18 14:54:53
Description  : file head
"""

from os import makedirs, system
from os.path import join, exists
from tkinter import Tk, filedialog

# import openpyxl
from pandas import concat

from alipay import ALiPBillWriter
from ccb import CCBBillWriter
from utils import BillWriter
from wechat import WCBillWriter


def starter(write_p, read_p: str = r"D:\Document\材料\账单&发票"):
    ccb_path = join(read_p, "建行")
    wx_path = join(read_p, "微信")
    ali_path = join(read_p, "支付宝")
    for i in [ccb_path, wx_path, ali_path]:
        if not exists(i):
            print("\n路径不存在\n")
            makedirs(i)
    ccb = CCBBillWriter(output_path=write_p)
    ccb_df = ccb.starter(ccb_path)
    wc = WCBillWriter(output_path=write_p)
    wc_df = wc.starter(wx_path)
    ali = ALiPBillWriter(output_path=write_p)
    ali_df = ali.starter(ali_path)
    final_df = concat([ccb_df, wc_df, ali_df])
    final_df.sort_values(by="交易时间", inplace=True)
    b = BillWriter(output_path=write_p)
    b.write(final_df)
    ali.refund_write()


if __name__ == '__main__':
    # path = r"C:\Users\Lenovo\OneDrive\Documents\个人财务管理.xlsx"
    config = "./config.txt"
    open(config, 'a')
    with open(config, 'r', encoding='utf-8') as file:
        lines = file.readlines()
        write_path = ""
        read_path = ""
        if len(lines) == 2:
            write_path = lines[0].strip()
            read_path = lines[1].strip()
            print(f"表格地址：{write_path}\n")
            print(f"账单地址：{read_path}\n")
        if (not exists(write_path)) or (not exists(read_path)):
            # 获取用户输入的字符串
            root = Tk()
            root.withdraw()
            if not exists(write_path):
                input("\n请选择记账文件地址（按回车键）\n")
                write_path = filedialog.askopenfilename(title="请选择记账表格（xlsx）",
                                                        filetypes=[("电子表格", ".xlsx")])
                print(f"表格地址：{write_path}\n")
            if not exists(read_path):
                input("\n请选择原始账单地址（按回车键）\n")
                read_path = filedialog.askdirectory(title="请选择原始账单地址")
                print(f"账单地址：{read_path}\n")
            open(config, 'w', encoding="utf-8").write(f"{write_path}\n{read_path}")
    starter(write_path, read_path)
    system(write_path)
    # system('chcp 65001')
    system('pause')
