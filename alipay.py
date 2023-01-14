from utils import BillReader, BillWriter


class ALiPBillReader(BillReader):
    __slots__ = ()

    def __init__(self, file=""):
        super().__init__(file)

    def read_file(self, filepath=""):
        super().read_file(filepath)


class ALiPBillWriter(BillWriter):
    __slots__ = ()

    def __init__(self, output_path=r"C:\Users\Lenovo\OneDrive\Documents\���˲������.xlsx", time_filter=True):
        super().__init__(output_path, time_filter)

    def starter(self, path=r"D:\\Document\\����\\�˵�&��Ʊ\\֧����\\"):
        for i in super().starter(path):
            print(i)
