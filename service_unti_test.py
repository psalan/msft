from my_unit_test import teszt
import msft_core_xlsl as t_core
import msft_research as t_rs
from msft_service_collect import Service_msft
import msft_patterns as pattern

# Fájl
file = './msft/files_xlsx/niexcel.xlsx'
core = t_core.Core_msft(file)

# Munkafüzet, munkalap
wb = core.wb
ws = wb.worksheets
core.ws_wide = 37

# első szakasz:
core.set_active_ws(0)
core.set_section_entry_point()
rs = t_rs.Research_msft(core)
core.active_research['ch_end_point'] = rs.rs_section_end_point()
rs.rs_is_end()
s = Service_msft(rs)


# adat koordináták megjelenítése
def view():
    """Simple adat cellák megmutatása."""
    # rs.rs_fix_it()
    # print('\n')
    # print(pattern.head_human_readable)
    # print('\n')
    # print(pattern.head)
    # print('\n')
    # print(pattern.data_head_human_readable)
    # print('\n')
    # print(pattern.data_head)

    # Egyéb tesztek
    print(s)
    head_coo = s.refresh_coordinates(s.change_dict2list(s.get_pattern_head()),
                                  s.head)
    print(s.get_content(head_coo))


def tesztkeszlet():
    """A modulhoz tartozó tesztkészlet futatása."""
    teszt(type(s.get_pattern_head()) == type(dict()))
    teszt(type(s.change_dict2list(s.get_pattern_head())) == type(list()))
    teszt(s.refresh_coordinates([(1, 5), (5, 2)], 5) != [(6, 5), (10, 2)])
    teszt(s.refresh_coordinates([(1, 5), (5, 2)], 5) == [(5, 5), (9, 2)])
    teszt(s.refresh_coordinates(s.change_dict2list(s.get_pattern_head()),\
        s.head) == [(1, 31), (5, 15), (7, 15), (9, 15), (11, 15)])


view()
tesztkeszlet()
