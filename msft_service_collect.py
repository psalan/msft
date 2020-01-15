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

    # def set_head(self, head_list: list, num: int):
    #     head = self.head
    #     temp = self.refresh_data(self.head - 1)
    #     for i in range(len(temp)):
    #         examine_cell = self.rs.rs_cell(self.ws, temp[i][0], temp[i][1])
    #         print(examine_cell.value)

    def set_content(self):
        pass
