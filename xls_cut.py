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

    def save_to_file(self):
        result = (self.check_line(self.sheet.row(i)[self.column].value) for i in range(self.sheet.nrows))
        with xlwt_context(self.file) as wb:
            ws = wb.add_sheet(self.sheet.name)
            for num, (f, s) in enumerate(result):
                ws.write(num, self.column, f)
                ws.write(num, self.column + 1, s)

    def check_line(self, line: str) -> namedtuple:
        res = self.format_line(line) if len(line) > 30 else FirstSecond(line, None)
        return res

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

    def format_line(self, line: str) -> namedtuple:
        counter = self.max_length
        my_line = line[:counter]
        while my_line[-1] != ' ':
            counter -= 1
            my_line = line[:counter]
        return FirstSecond(my_line, line[counter:])


@contextmanager
def xlwt_context(file):
    try:
        wb = xlwt.Workbook()
        yield wb
    finally:
        wb.save(f'MODIFED_{file}')


if __name__ == '__main__':
    XlsChanger()
