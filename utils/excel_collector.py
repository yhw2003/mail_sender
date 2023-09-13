import pandas as pd


class Data:
    name: str
    mail: str
    pass


class IterXlsx:
    def __init__(self, path: str):
        self._iter_index = 0
        self.raw_data = pd.read_excel(path)

    def __iter__(self):
        return self

    def __next__(self):
        raw_data = self.raw_data
        if self._iter_index < len(raw_data):
            _data = Data()
            _data.name = raw_data['姓名'][self._iter_index]
            _data.mail = raw_data['邮箱'][self._iter_index]
            self._iter_index += 1
            return _data
        raise StopIteration


if __name__ == "__main__":
    ds = IterXlsx("./test.xlsx")
    for data in ds:
        print(f"{data.name} {data.mail}")
