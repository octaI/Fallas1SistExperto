"""
Estructura del archivo de entrada

|Especie |Fecha	|Vencimiento| Tipo de Serie	|Cierre del día|Variación %	|Apertura |Máximo |Mínimo| Volumen Nominal|
Monto Negociado|N° Oper

"""

import pandas as pd
import xlrd


def load_database():
    excel_file = pd.ExcelFile("Merarg-input.xls")
    actions_sheet = excel_file.parse(0)
    return actions_sheet

def calculate_movingaverage(number,price_list):
    average_sum = 0
    list_length = len(price_list)
    starting_index = list_length -1
    for i in range(number):
        average_sum += price_list[starting_index - i]

    return average_sum/number


def calculate_ema(number, price_list):
    moving_average = calculate_movingaverage(number, price_list[:-1]) #stripping the relevant EMA's
    multiplier = (2/(number + 1))
    ema_to_return = (price_list[-1] - moving_average) * multiplier + moving_average

    return ema_to_return


def main():
    actions_sheet = load_database()


if __name__ == "__main__":
    main()
