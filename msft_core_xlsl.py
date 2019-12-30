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

        def __init__(self, entry_point=1):
            pass

    def core_auto_execution(self):
        """ Az excel fájl automatikus feldolgozási\
            folyamatát végző füügvény."""

        ftinf.get_str(ftinf.core['start'])
        ftinf.get_str_format(ftinf.core['ws_num'], len(self.ws_names))
