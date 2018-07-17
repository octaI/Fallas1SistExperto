"""
Estructura del archivo de entrada

|Especie |Fecha	|Vencimiento| Tipo de Serie	|Cierre del día|Variación %	|Apertura |Máximo |Mínimo| Volumen Nominal|
Monto Negociado|N° Oper

"""

import pandas as pd
import statistics
import sklearn


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
    return calculate_ema(12, price_list) - calculate_ema(26, price_list)

def calculate_PPO (price_list):
    return calculate_MACD(price_list) / calculate_ema(26,price_list) * 100

def calculate_RSI (price_list):
    relevant_numbers = price_list[-14:] #taking the last 14 days
    avg_gain = 0
    avg_loss = 0
    for num in relevant_numbers:
        if (num < 0): avg_gain+=num

        else:
            avg_loss+=(-num)
    avg_gain = avg_gain/14
    avg_loss = avg_loss/14
    relative_strength = avg_gain/avg_loss
    return (100 - (100/(1+relative_strength)))

def calculate_error_bands(price_list):
    regr = sklearn.linear_model.LinearRegression()
    x_values = list(range(1,31)) #taking the last 30 days for linear regression
    relevant_prices = price_list[-30]
    y_values = []
    for elem in relevant_prices:
        y_values.append(elem(0))
    regr.fit(x_values,y_values)
    std_dev = statistics.pstdev(y_values)
    return regr.predict(30) + 1.96*std_dev, regr.predict(30 - 1.96*std_dev)

def main():
    actions_sheet = load_database()
    relevant_data = actions_sheet[['Especie', 'Cierre del día', 'Variación %']]
    actions_dictionary = {}
    for index, row in relevant_data.iterrows():
        if row[0] not in actions_dictionary:
            actions_dictionary[row[0]] = [(row[1],row[2])]
        else:
            actions_dictionary[row[0]].append((row[1],row[2]))
    print(actions_dictionary['AGRO'])

if __name__ == "__main__":
    main()
