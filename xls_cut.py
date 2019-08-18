""" Для установки необходимых зависимостей выполнить команду:
pip install -r requirements.txt"""
import os
from contextlib import contextmanager
from collections import namedtuple
import xlrd
import xlwt

FirstSecond = namedtuple('FirstSecond', 'first second')


class XlsChanger:
    """xlsx_cut - скрипт, модифицирующий Excel файл"""

    def __init__(self):
        """Функция сначала принимает от пользователя максимальную длинну строки,
        затем номер файла, затем номер листа,
        затем номер колоки в листе"""
        self.max_length = int(input('\nМаксимальная строка: '))
        print()
        self.file = self.select_file()
        self.book = xlrd.open_workbook(self.file)
        self.sheet = self.select_sheet()
        self.column = int(input('\nВыберите колонку (А = 1, B = 2 и т.д.): ')) - 1
        print()
        self.save_to_file()

    def __del__(self):
        """Функция выводит сообщение о колличестве обработанных строк"""
        print(f'Обработано {len(self)} строк')

    def __getitem__(self, item) -> str:
        """Функция возвращает значение ячейки указанной колонке, в строке с заданным индексом"""
        return self.sheet.row(item)[self.column].value

    def __len__(self) -> int:
        """Функция возвращает колличество строк в файле"""
        return self.sheet.nrows

    def save_to_file(self):
        """Функция формирует генератор из значений всех ячеек, проверяя значнеия
        Затем, итератором передаёт по очереди значения на запись в файл"""
        result = (self.check_line(self.__getitem__(i), self.max_length) for i in range(len(self)))
        with xlwt_context(self.file) as write_file:
            write_sheet = write_file.add_sheet(self.sheet.name)
            for num, (first_value, second_value) in enumerate(result):
                write_sheet.write(num, self.column, first_value)
                write_sheet.write(num, self.column + 1, second_value)

    def select_sheet(self) -> xlrd.sheet.Sheet:
        """Функция выводит на экран все листы в файле и просит выбрать нужный"""
        for num, sheet in enumerate(self.book.sheet_names(), start=1):
            print(num, sheet)
        sheet_number = int(input('\nВыберите лист: ')) - 1
        return self.book.sheet_by_index(sheet_number)

    @staticmethod
    def select_file() -> str:
        """Функция выводит на экран все ФАЙЛЫ в дирректории и просит выбрать нужный"""
        only_files = (file for file in os.listdir(os.path.curdir)
                      if os.path.isfile(os.path.join(os.path.curdir, file)))
        for num, file in enumerate(only_files, start=1):
            print(num, file)
        file_number = int(input('\nВыберите файл: ')) - 1
        print()
        return os.listdir(os.path.curdir)[file_number]

    @staticmethod
    def check_line(line: str, max_length: int) -> namedtuple:
        """Функция выполняет проверку. Если длинна переданной ячейки больше максимальной длинны,
        вызывается функция очистки format_line, если меньше - создаётся именнованный кортеж формата
        (ЗНАЧЕНИЕ ЯЧЕЙКИ , None)"""
        res = XlsChanger.format_line(line, max_length) \
            if len(line) > max_length \
            else FirstSecond(line, None)
        return res

    @staticmethod
    def format_line(line: str, max_length: int) -> namedtuple:
        """Функция выполняет очистку. Данные ячейки обрезаются по ближайшему пробелу
        и возвращается именованный кортеж формата (ТЕКСТ ДО МАКСИМУМА СИМВОЛОВ , ОСТАЛЬНОЙ ТЕКСТ)
        Если пробела нет, а слово длиннее максимума - создаётся именнованный кортеж формата
        (None , ЗНАЧЕНИЕ ЯЧЕЙКИ)"""
        counter = max_length
        my_line = line[:counter]
        while my_line[-1] != ' ':
            counter -= 1
            my_line = line[:counter]
            if not my_line:
                return FirstSecond(None, line)
        return FirstSecond(my_line, line[counter:])


@contextmanager
def xlwt_context(file):
    """Менеджер контекста"""
    try:
        write_file = xlwt.Workbook()
        yield write_file
    finally:
        write_file.save(f'MODIFED_{file}')


if __name__ == '__main__':
    XlsChanger()
