"""
Estructura del archivo de entrada

|Especie |Fecha	|Vencimiento| Tipo de Serie	|Cierre del día|Variación %	|Apertura |Máximo |Mínimo| Volumen Nominal|
Monto Negociado|N° Oper

"""
from colorama import Fore,Style
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


def calculate_RSI(variation_list):
    relevant_numbers = variation_list[-14:]  # taking the last 14 days
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
    x_values = np.array(list(range(1, len(price_list)+1)))
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

def analyze_ma(ma,active_price):
    result = 0
    if ma > active_price:
        result = 1
    elif ma < active_price:
        result = -1

    return result

def analyze_macd(macd):
    result = 0
    if macd > 0:
        result = 1
    elif macd < 0:
        result = -1

    return result

def analyze_rsi(rsi):
    result = 0
    if (rsi >= 70):
        result = 1
    elif rsi <=30:
        result = -1
    return result

def analyze_ppo(ppo, prev_ppo):
    result = 0
    if (ppo > 0) and (prev_ppo < 0):
        result = -1
    elif (ppo < 0) and (prev_ppo > 0):
        result = 1

    return result

def analyze_errorbands(error_bands,active_price):
    result = 0
    upper_limit = error_bands[0]
    lower_limit = error_bands[1]
    if active_price >= upper_limit:
        result = 1
    elif active_price <= lower_limit:
        result = -1

    return result

def analyze_indicators(price_list,variations_list,action_name):
    action_score = 0
    ma5 = calculate_movingaverage(5,price_list)
    ma10 = calculate_movingaverage(10,price_list)
    ma20 = calculate_movingaverage(20,price_list)
    ma40 = calculate_movingaverage(40,price_list)
    ma100 = calculate_movingaverage(100,price_list)
    macd = calculate_MACD(price_list)
    ppo = calculate_PPO(price_list)
    prev_ppo = calculate_PPO(price_list[:-1])
    rsi = calculate_RSI(variations_list)
    error_bands = calculate_error_bands(price_list)
    active_price = price_list[-1]
    all_mas = [ma5,ma10,ma20,ma40,ma100]
    for ma in all_mas:
        action_score += analyze_ma(ma,active_price)

    action_score += analyze_macd(macd)
    action_score += analyze_rsi(rsi)
    action_score += analyze_ppo(ppo,prev_ppo)
    action_score += analyze_errorbands(error_bands,active_price)
    if action_score > 0:
        output_message = "VENDER"
        cond_colour = Fore.RED
    else:
        output_message = "COMPRAR"
        cond_colour = Fore.GREEN

    print(f"{Fore.BLUE}{Style.DIM}Habiendo obtenido los siguientes resultados sobre: {action_name} {Style.RESET_ALL}")
    print('')
    print('')
    print(f"{Fore.LIGHTRED_EX}{Style.DIM}********************************************************************"
          f"{Style.RESET_ALL}")
    print(f'{Fore.WHITE}{Style.BRIGHT} Valor: {active_price}  '
          f'MA5: {ma5}   MA10: {ma10}  MA20: {ma20}  MA40: {ma40}  MA100: {ma100}'
          f'  {Style.RESET_ALL}')
    print('')
    print(f'{Fore.WHITE}{Style.BRIGHT} MACD: {macd}  PPO: {ppo}  PPOprev: {prev_ppo}  RSI: {rsi}   {Style.RESET_ALL}')
    print('')
    print(f'{Fore.WHITE}{Style.BRIGHT} Limite Superior: {error_bands[0]}  Limite Inferior: {error_bands[1]} {Style.RESET_ALL}')
    print(f"{Fore.LIGHTRED_EX}{Style.DIM}********************************************************************"
          f"{Style.RESET_ALL}")
    print('')
    print('')
    print(f'{Fore.WHITE}{Style.BRIGHT}El sistema recomienda para {action_name}:'
          f' {cond_colour}{Style.BRIGHT}{output_message}{Style.RESET_ALL}')
    return

def main():
    print(f"{Fore.LIGHTRED_EX}{Style.DIM}*************************************{Style.RESET_ALL}")
    print(f"{Fore.LIGHTBLUE_EX}{Style.BRIGHT}BIENVENIDO AL ANALIZADOR DE ACCIONES{Style.RESET_ALL}")
    print(f"{Fore.LIGHTRED_EX}{Style.DIM}*************************************{Style.RESET_ALL}")
    print(f"{Fore.BLUE}{Style.DIM}Por favor ingrese la Especie que desea analizar{Style.RESET_ALL}")
    print(f"{Fore.BLUE}{Style.DIM}Usted tiene a disposicion las siguientes especies:{Style.RESET_ALL}")
    actions_dictionary = initial_load()
    names_df =pd.DataFrame({'Especie': list(actions_dictionary.keys())})
    with pd.option_context('display.max_rows',None,'display.max_columns',names_df.shape[1],'expand_frame_repr',True):
        print(f'{Fore.LIGHTYELLOW_EX}{names_df}{Style.RESET_ALL}')
    keep_running = True
    while keep_running:
        valid_name = False
        while (not valid_name):
            action_name = input("Accion: ")
            print('')
            print('')
            action_name = action_name.upper()
            if action_name.upper() not in actions_dictionary.keys():
                print(f"{Fore.LIGHTRED_EX}{Style.BRIGHT}Nombre de especie incorrecto. Por favor ingrese un nombre váli"
                      f"do{Style.RESET_ALL}")
            else:
                valid_name = True
        action_data = actions_dictionary[action_name] #getting the requested info
        variations_list = []
        prices_list = []
        for elem in action_data:
            variations_list.append(elem[1])
            prices_list.append(elem[0])  #splitting into the relevant lists

        analyze_indicators(prices_list,variations_list,action_name.upper())
        print('')
        print('')
        print('')
        print("Desea consultar otra Especie? Escriba Y en ese caso, cualquier otra para salir del sistema")
        response = input("--->")
        if response.upper() != 'Y':
            keep_running = False
    print('')
    print('')
    print('')
    print(f"{Fore.LIGHTRED_EX}{Style.DIM}*************************************{Style.RESET_ALL}")
    print (f'{Fore.LIGHTBLUE_EX}{Style.BRIGHT}GRACIAS POR UTILIZAR NUESTRO SISTEMA{Style.RESET_ALL}')
    print(f"{Fore.LIGHTRED_EX}{Style.DIM}*************************************{Style.RESET_ALL}")
if __name__ == "__main__":
    main()
