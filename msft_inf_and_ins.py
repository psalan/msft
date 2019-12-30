def get_str(keyword):
    """Megjeleníti a "keyword"-hőz tartozó szöveget."""
    print(keyword)

def get_str_format(*args):
    """Megjelenít egy formázandó szöveget. Az args[0]-nak kell lenni a\
        formázandó szöveg kulcsaának, mely egy dictionary-ben van\
        a többi args érték a format függvény argumentumai.\n
        args[0].format(args[1], args[2]...etc)->None"""
    word_list = list(args)
    keyword = word_list.pop(0)
    km_args = dict(enumerate(word_list))
    print(keyword.format(km_args))

frame = {'copy_file' : 'Másold az excel fájlt a "files_xlsx" könyvtárba!\
            \nMajd írd be a fájl nevét!',
        'answ_01' : "A fájl neve: ",  
        'answ_02' : 'Kilépsz a programból?\nIGEN\n',
        'Ok' : '\n\tRendben! Megtaláltam a fájlt.',
        'not_ok' : '\nA fájl nem található!',
        'bye' : 'Ok.\nVisszlát!'}

core = {'start' : '\n\tMegkezdem a fájl feldolgozásást.',
        'ws_num' : '\n\tA fájl {0[0]} darab munkalapot tartalmaz. {0[1]}'
        }