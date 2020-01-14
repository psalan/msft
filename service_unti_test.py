from my_unit_test import teszt
import msft_core_xlsl as t_core
import msft_research as t_rs
import msft_patterns as pattern

file = './msft/files_xlsx/niexcel.xlsx'
core = t_core.Core_msft(file)
wb = core.wb
ws = wb.worksheets
core.ws_wide = 37


# első szakaszI
def first_section():
    core.wipe_research()
    core.set_active_ws(0)
    core.set_section_entry_point()
    rs = t_rs.Research_msft(core)
    core.active_research['ch_end_point'] = rs.rs_section_end_point()
    rs.rs_is_end()


# másodki szakasz
def second_section():
    core.transmission_2store()
    core.set_section_entry_point()
    rs = t_rs.Research_msft(core)
    core.active_research['ch_end_point'] = rs.rs_section_end_point()
    rs.rs_is_end()


# adat koordináták megjelenítése
def view():
    first_section()
    rs = t_rs.Research_msft(core)
    rs.rs_fix_it()
    print('\n')
    print(pattern.head_human_readable)
    print('\n')
    print(pattern.head)
    print('\n')
    print(pattern.data_head_human_readable)
    print('\n')
    print(pattern.data_head)


def tesztkeszlet():
    pass


view()
tesztkeszlet()