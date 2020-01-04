import openpyxl
import msft_inf_and_ins as ftinf


class Core_msft:
    """Egy excel fájl feldolgozását segítő osztály."""

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

        # Aktív kutatás
        self.active_research = {'ws_poz': None,
                                'ws_name': None,
                                'ws_done': False,
                                'ch_entry_point': None,
                                'ch_end_point': 0,
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

        # sor hossza, ennyi oszlopból áll
        self.ws_wide_done = False
        self.ws_wide = None

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

    def rs_cell(self, worksheet, cell_row: int, cell_col: int):
        """ Visszatér egy meghatározott Cell objektummal.\n
        rs_cell(args)->Cell"""
        cell = worksheet.cell(cell_row, cell_col)
        return cell

    def count_cell(self, args_lista: list = [1, 1, 0, 2, ]):
        """ Visszatér a lista meghatározott cellájtól balra lévő, szintén\
            a lista által meghatározott egyesített cellához tartozó\
            oszlop számával. Kitétel, hogy a "merged_cell"-át,\
            "simple_cell"-nek kell követni.\n
            count_cell(list)->int"""
        # args_lista = [row_id, col_id, merged_cell_id, find_simple_cell]
        lista = args_lista
        examine_cell = self.rs_cell(
            self.ws[self.active_research['ws_poz']], lista[0], lista[1])
        if lista[3] == 0:
            return lista[2]
        else:
            if self.is_merged_cell(examine_cell):
                lista[2] = lista[1]
                lista[3] = 2    # általános kezdőérték
            if self.is_simple_cell(examine_cell):
                lista[3] -= 1
            lista[1] += 1
            return self.count_cell(lista)

    def set_section_entry_point(self):
        """ Beállítja az aktív kutatási szakasz belépési pontját.
            \n pre: elöző_kutatási_szakasz
            \n "belépési _pont" = "pre_kilépési_pont + 1 + \
            "pre_elválasztó_köz" """
        self.active_research['ch_entry_point'] = \
            self.previous_research['ch_end_point'] + \
            1 + self.previous_research['ch_separate']

    def rs_section_end_point(self):
        """ Megkeresi és beállítja az aktív kutatási szakasz\
            kilépési pontját."""
        empty_row = self.SEPARATE
        end_point = None
        row_id = self.active_research['ch_entry_point'] + 1
        while empty_row != 0:
            find_empty_row = self.count_cell([row_id, 1, 0, 2])
            if find_empty_row == 0:
                empty_row -= 1
                if empty_row == 0:
                    end_point = row_id - self.SEPARATE
            else:
                empty_row = self.SEPARATE
            row_id += 1
        self.active_research['ch_end_point'] = end_point

    def rs_is_end(self, num: int = 1):
        """ Ellenőrzi, hogy ez a szkasz volt-e az utolsó. Ha igen, akkor\
            beállítja az aktív kutatás értékét "True"-ra és visszaadja,\
            azt. Ellenkező esetben "False" értéket ad.\
            \nIlletve ellenőrzi a térköz állandóságát is.
            \nrs_is_end()->Boolean"""
        row_id = self.active_research['ch_end_point'] + self.SEPARATE + num
        examine_row = self.count_cell([row_id, 1, 0, 2])
        if examine_row == 0 and num == 1:
            num += 1
            return self.rs_is_end(num)
        elif examine_row == 0 and num != 1:
            self.active_research['ch_is_last'] = True
            return self.active_research['ch_is_last']
        elif examine_row != 0 and num == 1:
            self.active_research['ch_separate'] = self.SEPARATE
            return self.active_research['ch_is_last']
        else:
            ftinf.get_str_format(
                ftinf.core['check_file'], self.active_research['ch_end_point'])
            return exit()

    def transmission_2store(self):
        """ Az aktív tárolóból az adatok átvitele az ideiglenes tárolóba."""
        self.previous_research = self.active_research.copy()

    def wipe_research(self):
        """ Tárolók visszaállása a kiinduló helyzetbe. Új munkalap jön."""
        self.active_research = {'ws_poz': None,
                                'ws_name': None,
                                'ws_done': False,
                                'ch_entry_point': None,
                                'ch_end_point': None,
                                'ch_separate': 0,
                                'ch_is_last': False, }
        self.previous_research = {'ws_poz': None,
                                  'ws_name': None,
                                  'ws_done': False,
                                  'ch_entry_point': None,
                                  'ch_end_point': 0,
                                  'ch_separate': 0,
                                  'ch_is_last': None, }

    def core_main(self):
        """ Az excel fájl automatikus feldolgozását végző függvény."""

        ftinf.get_str(ftinf.core['start'])
        ftinf.get_str_format(ftinf.core['ws_num'], len(self.ws_names))
        for i in range(len(self.ws_names)):
            self.set_active_ws(i)
            while not self.active_research['ch_is_last']:
                if not self.ws_wide_done:
                    self.ws_wide = self.count_cell()
                    self.ws_wide_done = True
                self.set_section_entry_point()
                self.rs_section_end_point()
                self.rs_is_end()
                self.transmission_2store()
                print(self.active_research)
                print('Ide jön az adatgyűjtő rész.')
            self.wipe_research()

            # intra test (törlendő)
            # print(self.active_research)
            # print(self.rs_cell(
            #     self.ws[self.active_research['ws_poz']], 1, 1).coordinate)
