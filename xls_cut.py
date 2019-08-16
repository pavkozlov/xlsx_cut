import xlrd
import xlwt
import os


class XlsChanger:
    def __init__(self):
        self.max_length = int(input('Максимальная длинна строки: '))
        self.file = self.select_file()
        self.book = xlrd.open_workbook(self.file)
        self.sheet = self.select_sheet()
        self.column = int(input('Выберите колонку (А = 1, B = 2, C =3, D = 4, E = 5 и т.д.): ')) - 1
        self.write(self.save_to_csv())

    def save_to_csv(self):
        result = list()
        for i in range(self.sheet.nrows):
            line = self.sheet.row(i)[self.column].value
            res = self.format_line(line) if len(line) > 30 else {'f': line, 's': ''}
            result.append(res)
        return result

    def select_file(self):
        print(f'Выберите нужный файл из {len(os.listdir())}')
        for file in enumerate(os.listdir()):
            print(f'{file[0] + 1}) {file[1]}')
        file_number = int(input('Ответ: ')) - 1
        file = os.listdir()[file_number]
        print(f'Выбран файл: {file}')
        return file

    def select_sheet(self):
        print(f'Выберите нужный лист из {self.book.nsheets}: ')
        for sheet in enumerate(self.book.sheet_names()):
            print(f'{sheet[0] + 1}) {sheet[1]}')
        sheet_number = int(input('Ответ: ')) - 1
        sheet = self.book.sheet_by_index(sheet_number)
        print(f'Выбран лист: {sheet.name} ({sheet.nrows} строк * {sheet.ncols} колонок)')
        return sheet

    def write(self, result):
        wb = xlwt.Workbook()
        ws = wb.add_sheet(self.sheet.name)
        for row in enumerate(result):
            ws.write(row[0], self.column, row[1]['f'])
            ws.write(row[0], self.column + 1, row[1]['s'])
        wb.save(f'MODIFED_{self.file}')

    def format_line(self, line):
        counter = self.max_length
        my_line = line[:counter]
        while my_line[-1] != ' ':
            counter -= 1
            my_line = line[:counter]
        return {'f': my_line, 's': line[counter:]}


if __name__ == '__main__':
    XlsChanger()
