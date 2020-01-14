import msft_banner as ftbann
import msft_forda_teszt as ft
import msft_inf_and_ins as ftinf


class Frame_msft():
    """ Egy keretet ad az excel fájl teszteléséhez."""

    def __init__(self):
        # frame items
        self.dir_path = './msft/files_xlsx/'
        self.try_count = 0
        self.work_file = None

    def set_work_file(self, filename):
        """Beállítja a munkafájlt, az abszolút útvonalával."""
        self.work_file = filename

    def get_work_file(self):
        """Utasítást ad a leendő munka fájl elhelyezésére,\
            kéri annak elnevezését, azután hozzárendeli az\
            abszolút útvonalat. Visszaadja munkafájlt és az\
            útvonalát."""

        ftinf.get_str(ftinf.frame['copy_file'])
        work_file_name = input(ftinf.frame['answ_01'])
        work_file = self.dir_path + work_file_name
        return work_file

    def is_valid_filename(self, filename):
        """ Ellenőrzi, hogy a fájl létezik-e?"""
        try:
            with open(filename, "br") as f:
                f.read(1)
                ftinf.get_str(ftinf.frame['Ok'])
            return True
        except FileNotFoundError:
            ftinf.get_str(ftinf.frame['not_ok'])
            return False

    def app_end(self):
        """Kilép a programból."""
        print('kint')  # majd törölni
        exit()

    def is_end(self):
        """A kilépés lehetőségét ajánlja fel."""
        answer = str.lower(input(ftinf.frame['answ_02']))
        if not (answer == 'n' or answer == 'nem'):
            ftinf.get_str(ftinf.frame['bye'])
            self.app_end()

    def define_work(self):
        """Bekéri a munka fájlt, és gondoskodik az ellenőrzéséről."""
        done = False
        while True:
            if self.try_count > 1:
                self.is_end()
            filename = self.get_work_file()
            if self.is_valid_filename(filename):
                self.set_work_file(filename)
                done = True
                return done
            self.try_count += 1
            if self.try_count == 5:
                self.app_end()

    def start(self):
        """Elkezdi a keret a program felépítését."""
        if not (ftbann.make_banner(ftbann.banner)):
            self.app_end()
        else:
            if self.define_work():
                ft.Forda_teszt(self.work_file)

                self.app_end()
