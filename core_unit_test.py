from my_unit_test import teszt
import msft_core_xlsl as t_core
import msft_research as t_rs

file = './msft/files_xlsx/niexcel.xlsx'
core = t_core.Core_msft(file)
wb = core.wb
ws = wb.worksheets
core.ws_wide = 37
first_cell = ws[0].cell(1, 1)
second_cell = ws[0].cell(1, 2)


# első szakasz
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


def tesztkeszlet():
    teszt(len(core.ws) == 2)
    teszt(core.is_simple_cell(first_cell) == True)
    teszt(core.is_merged_cell(first_cell) == False)
    teszt(core.is_simple_cell(second_cell) == False)
    teszt(core.is_merged_cell(second_cell) == True)
    first_section()
    rs = t_rs.Research_msft(core)
    teszt(rs.rs_count_cell() != 35)
    teszt(rs.rs_count_cell() == 37)
    teszt(rs.rs_count_cell() != 38)
    teszt(core.active_research['ch_entry_point'] == 1)
    teszt(core.active_research['ch_end_point'] == 36)
    teszt(core.active_research['ch_end_point'] != 37)
    teszt(core.previous_research['ch_is_last'] == None)
    teszt(core.active_research['ch_is_last'] == False)
    second_section()
    teszt(core.active_research['ch_entry_point'] != 38)
    teszt(core.active_research['ch_entry_point'] == 39)
    teszt(core.active_research['ch_entry_point'] != 40)
    teszt(core.previous_research['ch_is_last'] == False)
    teszt(core.active_research['ch_is_last'] == True)


if __name__ == "__main__":
    tesztkeszlet()
