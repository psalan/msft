from msft_research import Research_msft
import msft_patterns as ftpatt


class Service_msft(Research_msft):
    """ Egy szolgálat adatait gyűjti ki."""

    def __init__(self, rs: Research_msft):
        self.rs = rs
        self.core = self.rs.core
        self.head = self.core.active_research['ch_entry_point']

    def __str__(self):
        """ Megjeleníti, hogy melyik fájl, munkalap, szakasz adatait\
            gyűjti össze."""
        serv_str = self.core.__str__() + ' munkalap' + '\n' + str(self.head)\
            + ' sortól - ' + str(self.core.active_research['ch_end_point'])\
            + ' sorig'
        return serv_str

    def get_pattern_head(self):
        head = ftpatt.head
        return head

    def refresh_data(self):
        key_list = list(self.get_pattern_head())
