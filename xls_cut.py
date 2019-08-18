import xlrd
import xlwt
import os
from contextlib import contextmanager
from collections import namedtuple

FirstSecond = namedtuple('FirstSecond', 'first second')


class XlsChanger:
    def __init__(self):
        self.max_length = int(input('Максимальная строка: '))
        self.file = self.select_file()
        self.book = xlrd.open_workbook(self.file)
        self.sheet = self.select_sheet()
        self.column = int(input('Выберите колонку (А = 1, B = 2 и т.д.): ')) - 1
        self.save_to_file()

    def __del__(self):
        print(f'Обработано {len(self)} строк')

    def __getitem__(self, item) -> str:
        return self.sheet.row(item)[self.column].value

    def __len__(self) -> int:
        return self.sheet.nrows

    def save_to_file(self):
        result = (self.check_line(self.__getitem__(i), self.max_length) for i in range(len(self)))
        with xlwt_context(self.file) as wb:
            ws = wb.add_sheet(self.sheet.name)
            for num, (f, s) in enumerate(result):
                ws.write(num, self.column, f)
                ws.write(num, self.column + 1, s)

    def select_file(self) -> str:
        print(f'Выберите файл:')
        for num, file in enumerate(os.listdir(), start=1):
            print(num, file)
        file_number = int(input('Ответ: ')) - 1
        return os.listdir()[file_number]

    def select_sheet(self) -> xlrd.sheet.Sheet:
        print(f'Выберите лист: ')
        for num, sheet in enumerate(self.book.sheet_names(), start=1):
            print(num, sheet)
        sheet_number = int(input('Ответ: ')) - 1
        return self.book.sheet_by_index(sheet_number)

    @staticmethod
    def check_line(line: str, max_length: int) -> namedtuple:
        res = XlsChanger.format_line(line, max_length) if len(line) > max_length else FirstSecond(line, None)
        return res

    @staticmethod
    def format_line(line: str, max_length: int) -> namedtuple:
        if len(line) > max_length:
            counter = max_length
            my_line = line[:counter]
            while my_line[-1] != ' ':
                counter -= 1
                my_line = line[:counter]
                if my_line == '':
                    return FirstSecond(None, line)
            return FirstSecond(my_line, line[counter:])
        else:
            return FirstSecond(line, None)


@contextmanager
def xlwt_context(file):
    try:
        wb = xlwt.Workbook()
        yield wb
    finally:
        wb.save(f'MODIFED_{file}')


if __name__ == '__main__':
    XlsChanger()
