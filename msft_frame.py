import msft_banner as ftbann


# frame items
dir_path = './files_xlsx/'
try_count = 0


def get_work_file():
    print('Másold az excel fájlt a "files_xlsx" könyvtárba!\
        \nMajd írd be a fájl nevét!')

    work_file = input("A fájl neve: ")
    work_file = dir_path + work_file
    return work_file


def is_valid_filename(file):
    """ Ellenőrzi, hogy a fájl létezik-e?"""
    try:
        with open(file, "br") as f:
            f.read(1)
            print('\n\tRendben! Megtaláltam a fájlt.')
        return True
    except FileNotFoundError:
        print('\nA fájl nem található!')
        return False


def app_end():
    print('kint')
    exit()


def is_end():
    answer = str.lower(input('Kilépsz a programból?\
        \nIGEN\n'))
    if not (answer == 'n' or answer == 'nem'):
        print('Ok.\nVisszlát!')
        app_end()


if __name__ == "__main__":
    ftbann.make_banner(ftbann.banner)

    while True:
        if try_count > 1:
            is_end()
        filename = get_work_file()
        if is_valid_filename(filename):
            break
        try_count += 1

    print('Itt lesz a "mag"')
