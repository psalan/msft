from msft_core_xlsl import Core_msft
import msft_inf_and_ins as ftinf
import msft_patterns as ftpatt


class Research_msft(Core_msft):
    """ Ez az osztály kereső függvényeket valósít meg a Core osztály\
        részére."""

    def __init__(self, core: Core_msft):
        """ Megadjuk neki, hogy melyik példányon,\
            hajtjuk végre a kereséséeket."""
        self.core = core

    def rs_cell(self, worksheet, cell_row: int, cell_col: int):
        """ Visszatér egy meghatározott Cell objektummal.\n
        rs_cell(args)->Cell"""
        cell = worksheet.cell(cell_row, cell_col)
        return cell

    def rs_count_cell(self, args_lista: list = [1, 1, 0, 2, ]):
        """ Visszatér a lista meghatározott cellájtól balra lévő, szintén\
            a lista által meghatározott egyesített cellához tartozó\
            oszlop számával. Kitétel, hogy a "merged_cell"-át,\
            "simple_cell"-nek kell követni.\n
            count_cell(list)->int"""
        # args_lista = [row_id, col_id, merged_cell_id, find_simple_cell]
        lista = args_lista
        examine_cell = self.rs_cell(
            self.core.ws[self.core.active_research['ws_poz']], lista[0], lista[1])
        if lista[3] == 0:
            return lista[2]
        else:
            if self.core.is_merged_cell(examine_cell):
                lista[2] = lista[1]
                lista[3] = 2    # általános kezdőérték
            if self.core.is_simple_cell(examine_cell):
                lista[3] -= 1
            lista[1] += 1
            return self.rs_count_cell(lista)

    def rs_section_end_point(self):
        """ Visszaadja az aktív kutatási szakasz kilépési pontját.\
            \nrs_section_end_point()->Int"""
        empty_row = self.core.SEPARATE
        end_point = None
        row_id = self.core.active_research['ch_entry_point'] + 1
        while empty_row != 0:
            find_empty_row = self.rs_count_cell([row_id, 1, 0, 2])
            if find_empty_row == 0:
                empty_row -= 1
                if empty_row == 0:
                    end_point = row_id - self.core.SEPARATE
            else:
                empty_row = self.core.SEPARATE
            row_id += 1
        return end_point

    def rs_is_end(self, num: int = 1):
        """ Ellenőrzi, hogy ez a szkasz volt-e az utolsó. Ha igen, akkor\
            beállítja az aktív kutatás értékét "True"-ra és visszaadja,\
            azt. Ellenkező esetben "False" értéket ad.\
            \nIlletve ellenőrzi a térköz állandóságát is.
            \nrs_is_end()->Boolean"""
        row_id = self.core.active_research['ch_end_point'] + \
            self.core.SEPARATE + num
        examine_row = self.rs_count_cell([row_id, 1, 0, 2])
        if examine_row == 0 and num == 1:
            num += 1
            return self.rs_is_end(num)
        elif examine_row == 0 and num != 1:
            self.core.active_research['ch_is_last'] = True
            return self.core.active_research['ch_is_last']
        elif examine_row != 0 and num == 1:
            self.core.active_research['ch_separate'] = self.core.SEPARATE
            return self.core.active_research['ch_is_last']
        else:
            ftinf.get_str_format(
                ftinf.core['check_file'], self.core.active_research['ch_end_point'])
            return exit()

    def rs_data_cell_positions(self, n_from: int, n_to: int):
        """ Egy tartományon belüli "simple cell" koordinátáit (excel és\
            sor/oszlop szám formátumban) és értékeit adja vissza listákba\
            rendezve.\
            \nrs_data_cell_positions(*args)->Tuple"""
        examine_range = n_to + 1
        cell_coordinates = []
        cell_ids = []
        cell_values = []
        for n_from in range(n_from, examine_range,):
            for i in range(1, self.core.ws_wide + 1):
                examine_cell = self.rs_cell(
                    self.core.ws[self.core.active_research['ws_poz']], n_from, i)
                if self.core.is_simple_cell(examine_cell):
                    cell_coordinates.append(examine_cell.coordinate)
                    ids = (n_from, i)
                    cell_ids.append(ids)
                    cell_values.append(examine_cell.value)
        return (cell_coordinates, cell_ids, cell_values)

    def rs_fix_it(self, head: tuple = (1, 12), data: tuple = (14, 15)):
        """ Új, valós értékekre állítja be az adatgyűjtés koordinátáit."""
        head_cells = self.rs_data_cell_positions(head[0], head[1])
        ftpatt.head = dict(zip(head_cells[0], head_cells[1]))
        ftpatt.head_human_readable = dict(zip(head_cells[0], head_cells[2]))
        data_cells = self.rs_data_cell_positions(data[0], data[1])
        ftpatt.data_head = dict(zip(data_cells[0], data_cells[1]))
        ftpatt.data_head_human_readable = dict(
            zip(data_cells[0], data_cells[2]))
