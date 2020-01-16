from msft_research import Research_msft
import msft_patterns as ftpatt


class Service_msft(Research_msft):
    """ Egy szolgálat adatait gyűjti ki."""

    def __init__(self, rs: Research_msft):
        self.rs = rs
        self.core = self.rs.core
        self.ws = self.core.ws[self.core.active_research['ws_poz']]
        self.head = self.core.active_research['ch_entry_point']

    def __str__(self):
        """ Megjeleníti, hogy melyik fájl, munkalap, szakasz adatait\
            gyűjti össze."""
        serv_str = self.core.__str__() + ' munkalap' + '\n' + str(self.head)\
            + ' sortól - ' + str(self.core.active_research['ch_end_point'])\
            + ' sorig'
        return serv_str

    def get_pattern_head(self):
        """ Visszaad egy dictionary-ben tárolt sablont, ez a fejléc kijelölt\
            adatainak koordinátáit tartalmazza.\
            \nget_pattern_head()->dict"""
        head = ftpatt.head
        return head

    def change_dict2list(self, t_dict: dict):
        """ A cél könyvtárban (t_dict) átadott key:value párok értékeit egy\
            listában adja vissza.\
            \nrefresh_data(dict)->list"""
        key_list = list(t_dict)
        value_list = []
        for i in key_list:
            value_list.append(t_dict[i])
        return value_list

    def refresh_coordinates(self, coo_list: list, num: int):
        """ A coo_list-ában szereplő valamennyi koordinátának a sor értékét\
            megnöveli num értékével. Majd (a tuple-k miatt) egy új listában\
            adja vissza az új koordinátákat.\
            \nset_coordinates(*args)->list"""
        new_coo_list = []
        for i in range(len(coo_list)):
            coordinate = coo_list[i]
            row_id = coordinate[0]
            new_row_id = row_id + num - 1
            new_coordinate = (new_row_id, coordinate[1])
            new_coo_list.append(new_coordinate)
        return new_coo_list

    def get_content(self, coo_list: list):
        """ Visszaad egy listát, amely a coo_list-ben levő koordinátapárokon\
            található cellák értékeit tartalmazza\
            \nget_content(coo_list)->list"""
        values = []
        for i in range(len(coo_list)):
            coordinate = coo_list[i]
            e_cell = self.rs.rs_cell(self.ws, coordinate[0], coordinate[1])
            values.append(e_cell.value)
        return values
