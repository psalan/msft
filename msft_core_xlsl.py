import openpyxl
import msft_inf_and_ins as ftinf


class Core_msft:
    """Egy excel fájl feldolgozását segítő osztály."""

    def __init__(self, filename):

        # Konstansok
        self.separate = 3
        self.SIMPLE_CELL = "<class 'openpyxl.cell.cell.Cell'>"
        self.MERGED_CELL = "<class 'openpyxl.cell.cell.MergedCell'>"

        # Munkafájl
        self.work_file = filename

        # Munkafüzet
        self.wb = openpyxl.load_workbook(self.work_file)

        # Munkalapok
        self.ws = self.wb.worksheets

        # Munkalapok neve
        self.ws_names = self.wb.sheetnames

        # Aktív kutatás
        self.active_research = {'ws_poz': None,
                                'ws_name': None,
                                'ws_done': False,
                                'ch_entry_point': None, 
                                'ch_end_point': None,
                                'ch_is_last': None,}

        # Elöző kutatás
        self.previous_research = {'ws_poz': None,
                                'ws_name': None,
                                'ws_done': False,
                                'ch_entry_point': None, 
                                'ch_end_point': 0,
                                'ch_is_last': None,}                                

        # sor hossza, ennyi oszlopból áll
        #

    def is_simple_cell(self, cell):
        if str(type(cell)) == self.SIMPLE_CELL:
            return True
        return False

    def is_merged_cell(self, cell):
        if str(type(cell)) == self.MERGED_CELL:
            return True
        return False

    def set_active_ws(self, num: int):
        """ Beállítja az aktív munkalap értékeit. """
        self.active_research['ws_poz'] = num
        self.active_research['ws_name'] = self.ws_names[num]


    def ws_wide(self):
        """ Visszatér a munkalap egy adott sorának utolsó\
            egyesített cellájához tartozó oszlop számával. """
        count = 1
        col_id = 1
        find_simple_cell = self.separate
        # addig keres, amíg 3db nem egyesített cellát talál egymás után
        while find_simple_cell != 0:
            examine_cell = self.ws[0].cell(1, count)
            if self.is_simple_cell(examine_cell):
                find_simple_cell -= 1
            if self.is_merged_cell(examine_cell):
                col_id = count
                find_simple_cell = self.separate
            count += 1
        return col_id

    # Kutató alosztály
    class Research_msft:
        """ A excel munkalapon a szakaszokra bontott\
            Szolgálatok kezdő és végpontjait kutatja fel. """

        def __init__(self, active_research: set):
            """ A szolgálatok leírását tartalmazó szakaszok keresését\
                valósítja meg. """
            pass

    def core_main(self):
        """ Az excel fájl automatikus feldolgozását végző függvény."""

        ftinf.get_str(ftinf.core['start'])
        ftinf.get_str_format(ftinf.core['ws_num'], len(self.ws_names))
        for i in range(len(self.ws_names)):
            self.set_active_ws(i)

            print(self.active_research)
