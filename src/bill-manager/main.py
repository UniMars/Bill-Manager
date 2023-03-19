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
import json
# import openpyxl
from pandas import concat, DataFrame

from utils.alipay import ALiPBillWriter
from utils.ccb import CCBBillWriter
from utils.util import BillWriter
from utils.wechat import WCBillWriter


def writer(write_p, read_p: str = r"D:\Document\材料\账单&发票", is_wechat=True, is_alipay=True, is_ccb=True):
    ccb_path = join(read_p, "建行")
    wx_path = join(read_p, "微信")
    ali_path = join(read_p, "支付宝")
    final_df = DataFrame()
    for i in [ccb_path, wx_path, ali_path]:
        if not exists(i):
            print("\n路径不存在\n")
            makedirs(i)
    if is_wechat:
        if not exists(wx_path):
            makedirs(i)
        wc = WCBillWriter(output_path=write_p)
        wc_df = wc.starter(wx_path)
        final_df = concat([final_df, wc_df])
    ali = ALiPBillWriter(output_path=write_p)
    if is_alipay:
        if not exists(ali_path):
            makedirs(ali_path)
        ali_df = ali.starter(ali_path)
        final_df = concat([final_df, ali_df])
    if is_ccb:
        if not exists(ccb_path):
            makedirs(ccb_path)
        ccb = CCBBillWriter(output_path=write_p)
        ccb_df = ccb.starter(ccb_path)
        final_df = concat([final_df, ccb_df])
    if len(final_df):
        final_df.sort_values(by="交易时间", inplace=True)
    b = BillWriter(output_path=write_p)
    b.write(final_df)
    if is_alipay:
        ali.refund_write()


def starter():
    config_path = "./config.json"
    open(config_path, 'a')
    with open(config_path, 'r', encoding='utf-8') as file:
        try:
            config = json.load(file)
        except json.decoder.JSONDecodeError:
            config = {}
        write_path = config.get('write_path')
        read_path = config.get('read_path')
        is_wechat = bool(config.get("wechat"))
        is_alipay = bool(config.get("alipay"))
        is_ccb = bool(config.get("ccb"))
        write_path = "" if (not write_path) else write_path
        read_path = "" if (not read_path) else read_path

        print(f"\n表格地址：{write_path}\n")
        print(f"\n账单地址：{read_path}\n")

        if (not exists(write_path)) or (not exists(read_path)):
            if not config:
                is_alipay = True
                is_wechat = True
                is_ccb = input("\n是否导入建行账单（是请输入y后按回车键）\n") in ['y', "Y", '']
            # 获取用户输入的字符串
            root = Tk()
            root.withdraw()
            if not exists(write_path):
                input("\n请选择记账文件地址（按回车键）\n")
                write_path = filedialog.askopenfilename(title="请选择记账表格（xlsx）",
                                                        filetypes=[("电子表格", ".xlsx")])
                config['write_path'] = write_path
                print(f"表格地址：{write_path}\n")
            if not exists(read_path):
                input("\n请选择原始账单地址（按回车键）\n")
                read_path = filedialog.askdirectory(title="请选择原始账单地址")
                config['read_path'] = read_path
                print(f"账单地址：{read_path}\n")

        config['wechat'] = is_wechat
        config['alipay'] = is_alipay
        config['ccb'] = is_ccb
        open(config_path, 'w', encoding="utf-8").write(json.dumps(config, ensure_ascii=False))

    writer(write_path, read_path, is_wechat, is_alipay, is_ccb)
    system(write_path)


if __name__ == '__main__':
    starter()
    # system('chcp 65001')
    system('pause')
