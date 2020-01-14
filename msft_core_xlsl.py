import openpyxl


class Core_msft:
    """Egy excel fájl feldolgozását segítő alap osztály."""

    def __init__(self, filename):

        # Konstansok
        self.SIMPLE_CELL = "<class 'openpyxl.cell.cell.Cell'>"
        self.MERGED_CELL = "<class 'openpyxl.cell.cell.MergedCell'>"
        self.SEPARATE = 2

        # Munkafájl
        self.work_file = filename

        # Munkafüzet
        self.wb = openpyxl.load_workbook(self.work_file)

        # Munkalapok
        self.ws = self.wb.worksheets

        # Munkalapok neve
        self.ws_names = self.wb.sheetnames

        # Tárolók beállítása
        self.wipe_research()

        # sor hossza, ennyi oszlopból áll
        self.ws_wide_done = False
        self.ws_wide = None

    def __str__(self):
        core_str = self.work_file + '\n' + self.active_research['ws_name']
        return core_str

    def is_simple_cell(self, cell):
        """ Ha az argumentum cella típusa "simple', akkor True\
            értéket ad vissza, minden más esetben False-t\n
            is_simple_cell(cell)->Boolean"""
        if str(type(cell)) == self.SIMPLE_CELL:
            return True
        return False

    def is_merged_cell(self, cell):
        """ Ha az argumentum cella típusa "merged', akkor True\
            értéket ad vissza, minden más esetben False-t\n
            is_merged_cell(cell)->Boolean"""
        if str(type(cell)) == self.MERGED_CELL:
            return True
        return False

    def set_active_ws(self, num: int):
        """ Beállítja az aktív munkalap értékeit. """
        self.active_research['ws_poz'] = num
        self.active_research['ws_name'] = self.ws_names[num]

    def set_section_entry_point(self):
        """ Beállítja az aktív kutatási szakasz belépési pontját.
            \n pre: elöző_kutatási_szakasz
            \n "belépési _pont" = "pre_kilépési_pont + 1 + \
            "pre_elválasztó_köz" """
        self.active_research['ch_entry_point'] = \
            self.previous_research['ch_end_point'] + \
            1 + self.previous_research['ch_separate']

    def transmission_2store(self):
        """ Az aktív tárolóból az adatok átvitele az ideiglenes tárolóba."""
        self.previous_research = self.active_research.copy()

    def wipe_research(self):
        """ Tárolók visszaállása a kiinduló helyzetbe. Új munkalap jön."""
        # Aktív kutatás
        self.active_research = {'ws_poz': None,
                                'ws_name': None,
                                'ws_done': False,
                                'ch_entry_point': None,
                                'ch_end_point': None,
                                'ch_separate': 0,
                                'ch_is_last': False, }
        # Elöző kutatás
        self.previous_research = {'ws_poz': None,
                                  'ws_name': None,
                                  'ws_done': False,
                                  'ch_entry_point': None,
                                  'ch_end_point': 0,
                                  'ch_separate': 0,
                                  'ch_is_last': None, }
