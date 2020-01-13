import msft_inf_and_ins as ftinf
from msft_core_xlsl import Core_msft
from msft_research import Research_msft

class Forda_teszt:
    """ Végre hajtja forda tesztelését."""

    def __init__(self, work_file):
        """ Az alapbeállítások után meghivja a saját main() metódusát."""
        self.work_file = work_file
        self.main()

    def main(self):
        """ Az excel fájl automatikus feldolgozását végző függvény."""

        ftinf.get_str(ftinf.core['start'])
        base = Core_msft(self.work_file)
        ftinf.get_str_format(ftinf.core['ws_num'], len(base.ws_names))
        for i in range(len(base.ws_names)):
            base.set_active_ws(i)
            rs = Research_msft(base)
            while not base.active_research['ch_is_last']:
                if not base.ws_wide_done:
                    base.ws_wide = rs.rs_count_cell()
                    base.ws_wide_done = True
                base.set_section_entry_point()
                rs.rs_section_end_point()
                rs.rs_is_end()
                base.transmission_2store()
                print(base.active_research)
                print('Ide jön az adatgyűjtő rész.')
            base.wipe_research()

