"""
Estructura del archivo de entrada

|Especie |Fecha	|Vencimiento| Tipo de Serie	|Cierre del día|Variación %	|Apertura |Máximo |Mínimo| Volumen Nominal|
Monto Negociado|N° Oper

"""
import numpy as np
import pandas as pd
import statistics
from sklearn import linear_model

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


def calculate_PPO(price_list):
    return (calculate_MACD(price_list) / calculate_ema(26, price_list)) * 100


def calculate_RSI(price_list):
    relevant_numbers = price_list[-14:]  # taking the last 14 days
    avg_gain = 0
    avg_loss = 0
    for num in relevant_numbers:
        if num < 0:
            avg_gain += num

        else:
            avg_loss += (-num)
    avg_gain = avg_gain / 14
    avg_loss = avg_loss / 14
    relative_strength = avg_gain / avg_loss
    return 100 - (100 / (1 + relative_strength))


def calculate_error_bands(price_list):
    regr = linear_model.LinearRegression()
    x_values = np.array(list(range(1, len(price_list)+1)))  # taking the last 30 days for linear regression
    x_values = x_values.reshape(-1, 1)
    y_values = np.array(price_list)
    y_values = y_values.reshape(-1,1)
    regr.fit(x_values, y_values)
    std_dev = statistics.pstdev(price_list) #std deviation of the price list
    desired_point = np.array([len(price_list)]).reshape(-1,1)
    return [regr.predict(desired_point)[0,0] + 1.96*std_dev,regr.predict(desired_point)[0,0]-1.96*std_dev]

def initial_load():
    actions_sheet = load_database()
    relevant_data = actions_sheet[['Especie', 'Cierre del día', 'Variación %']]
    actions_dictionary = {}
    for index, row in relevant_data.iterrows():
        if row[0] not in actions_dictionary:
            actions_dictionary[row[0]] = [(row[1], row[2])]
        else:
            actions_dictionary[row[0]].append((row[1], row[2]))

    return actions_dictionary


def analyze_indicators(price_list):

    return

def main():
    actions_dictionary = initial_load()
    print("*************************************")
    print("BIENVENIDO AL ANALIZADOR DE ACCIONES")
    print("*************************************")
    print("Por favor ingrese la Especie que desea analizar")
    print("Usted tiene a disposicion las siguientes especies:")
    names_df =pd.DataFrame({'Especie': list(actions_dictionary.keys())})
    with pd.option_context('display.max_rows',None,'display.max_columns',names_df.shape[1],'expand_frame_repr',True):
        print(names_df)

    valid_name = False
    while (not valid_name):
        action_name = input("Accion: ")
        action_name = action_name.upper()
        if action_name.upper() not in actions_dictionary.keys():
            print("Nombre de especie incorrecto. Por favor ingrese un nombre válido")
        else:
            print("Usted ha elegido "+action_name)
            valid_name = True
    action_data = actions_dictionary[action_name] #getting the requested info
    variations_list = []
    prices_list = []
    for elem in action_data:
        variations_list.append(elem[1])
        prices_list.append(elem[0])  #splitting into the relevant lists

if __name__ == "__main__":
    main()
