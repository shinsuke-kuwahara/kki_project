from kki_function.blynk import pins_read
from kki_function.blynk import pin_write
from kki_function.blynk import pins_clear
from kki_function.blynk import lcd_write

from kki_function.blynk import lcd_write
from kki_function.blynk import pins_read
from kki_function.blynk import pin_write

PIN = 'V9'
PIN1 = 'V10'

PINS = [PIN, PIN1]

TOKEN = "FJONUUiXQb7dQafFyi_GtFXc7baNdQpF"


def main():

    while True:
        buttons = pins_read(TOKEN, PINS)

        print(buttons)


if __name__ == '__main__':
    main()