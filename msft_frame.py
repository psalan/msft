import msft_banner as ftbann


class Frame_msft():

    def __init__(self):
        # frame items
        self.dir_path = './msft/files_xlsx/'
        self.try_count = 0

    def get_work_file(self):
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
        print('kint') # majd törölni
        exit()

    def is_end(self):
        answer = str.lower(input('Kilépsz a programból?\
            \nIGEN\n'))
        if not (answer == 'n' or answer == 'nem'):
            print('Ok.\nVisszlát!')
            self.app_end()

    def start(self):
        """Elkezdi a keret a program felépítését."""
        if (ftbann.make_banner(ftbann.banner)):

            while True:
                if self.try_count > 1:
                    self.is_end()
                filename = self.get_work_file()
                if self.is_valid_filename(filename):
                    break
                self.try_count += 1

            print('Itt lesz a "mag"')

        else:
            self.app_end()


if __name__ == "__main__":
    starter = Frame_msft()
    starter.start()
