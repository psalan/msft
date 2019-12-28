import msft_banner as ftbann


class Frame_msft():

    def __init__(self):
        # frame items
        self.dir_path = './msft/files_xlsx/'
        self.try_count = 0

    def get_work_file(self):
        """Utasítást ad a munka fájl elhelyezésére,\
            visszaadja munkafájlt és az útvonalát."""

        print('Másold az excel fájlt a "files_xlsx" könyvtárba!\
            \nMajd írd be a fájl nevét!')

        work_file = input("A fájl neve: ")
        work_file = self.dir_path + work_file
        return work_file

    def is_valid_filename(self, file):
        """ Ellenőrzi, hogy a fájl létezik-e?"""
        try:
            with open(file, "br") as f:
                f.read(1)
                print('\n\tRendben! Megtaláltam a fájlt.')
            return True
        except FileNotFoundError:
            print('\nA fájl nem található!')
            return False

    def app_end(self):
        """Kilép a programból."""
        print('kint')  # majd törölni
        exit()

    def is_end(self):
        """A kilépés lehetőségét ajánlja fel."""
        answer = str.lower(input('Kilépsz a programból?\
            \nIGEN\n'))
        if not (answer == 'n' or answer == 'nem'):
            print('Ok.\nVisszlát!')
            self.app_end()

    def define_work(self):
        """Bekéri a munka fájlt, és gondoskodik az ellenőrzéséről."""
        done = False
        while True:
            if self.try_count > 1:
                self.is_end()
            filename = self.get_work_file()
            if self.is_valid_filename(filename):
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
                print('Itt lesz a "mag"')
                self.app_end()
