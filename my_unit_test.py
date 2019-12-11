import sys


def teszt(sikeres_teszt):
    """Egy teszt eredményének a megjelenítése."""
    sorszam = sys._getframe(1).f_lineno
    try:
        assert sikeres_teszt
        print(f'A {sorszam}. teszt sikeres.')
    except AssertionError:
        print(f'A {sorszam}. teszt SIKERTELEN!!!')


def tesztkeszlet():
    """A modulhoz tartozó tesztkészlet futatása."""
    teszt(1 == 1)
    teszt(2 != 3)
    teszt(4 == 5)


if __name__ == "__main__":
    tesztkeszlet()
