from my_unit_test import teszt
import msft_core_xlsl as t_core

file = './msft/files_xlsx/niexcel.xlsx'
core = t_core.Core_msft(file)
wb = core.wb
ws = wb.worksheets
first_cell = ws[0].cell(1, 1)
second_cell = ws[0].cell(1, 2)


def tesztkeszlet():
    teszt(len(core.ws) == 2)
    teszt(core.is_simple_cell(first_cell) == True)
    teszt(core.is_merged_cell(first_cell) == False)
    teszt(core.is_simple_cell(second_cell) == False)
    teszt(core.is_merged_cell(second_cell) == True)
    teszt(core.ws_wide() == 37)

if __name__ == "__main__":
    tesztkeszlet()