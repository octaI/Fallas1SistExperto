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


def calculate_movingaverage(number, price_list):
    average_sum = 0
    list_length = len(price_list)
    starting_index = list_length - 1
    for i in range(number):
        average_sum += price_list[starting_index - i]

    return average_sum / number


def calculate_ema(number, price_list):
    moving_average = calculate_movingaverage(number, price_list[:-number])  # stripping the relevant EMA's
    multiplier = (2 / (number + 1))
    initial_ema = (price_list[-number] - moving_average) * multiplier + moving_average
    for i in reversed(range(number)):
        new_ema = (price_list[-i] - initial_ema) * multiplier + initial_ema
        initial_ema = new_ema

    return new_ema


def calculate_MACD(price_list):
    return calculate_ema(26, price_list) - calculate_ema(13, price_list)


def calculate_signalline(price_list):
    return calculate_ema(9, price_list)


def main():
    actions_sheet = load_database()
    relevant_data = actions_sheet[['Especie', 'Cierre del día']]
    actions_dictionary = {}
    for index, row in relevant_data.iterrows():
        if row[0] not in actions_dictionary:
            actions_dictionary[row[0]] = [row[1]]
        else:
            actions_dictionary[row[0]].append(row[1])


if __name__ == "__main__":
    main()
