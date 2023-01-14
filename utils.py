import os
import shutil

import openpyxl
import pandas as pd


class BillReader(object):
    __slots__ = '_file', "_type", "_host_name", "_output_path", "_time_filter"

    def __init__(self, file=""):
        self._type = ""
        self._file = file
        # if file != "":
        #     self.read_file(file)

    def read_file(self, filepath=""):
        assert filepath != "" or self._file != "", "no input location"
        filepath = self._file if filepath == "" else filepath
        self._file = filepath
        suffix = filepath.split('.')[-1]
        assert suffix == "csv" or suffix == "xlsx" or suffix == "xls", "file type is not sheet"
        self._type = suffix
        print(f"type:{suffix}\n")


class BillWriter(object):
    __slots__ = "_output_path", "_time_filter"

    def __init__(self, output_path=r"C:\Users\Lenovo\OneDrive\Documents\个人财务管理.xlsx", time_filter=True):
        self._output_path = output_path
        self._time_filter = time_filter

    def starter(self, path):
        # df_list = []
        for i in os.listdir(path):
            if i == "old":
                continue
            yield i
            shutil.move(path + i, path + "old//")

    def write(self, df: pd.DataFrame):
        book = openpyxl.load_workbook(self._output_path)
        # a=book.get_sheet_names()
        sh = book.get_sheet_by_name('流水表')
        t = 0
        if self._time_filter:
            t = list(sh.rows)[-1][0].value
        if t and t != "交易时间":
            df = df[df['交易时间'] > t]
        return book, df, sh
