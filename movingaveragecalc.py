import argparse
import configparser
import gspread
import numpy as np
from oauth2client.service_account import ServiceAccountCredentials


def get_data(credentials_json, ssheet_id, scope):
    try:
        credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_json, scope)
        gs = gspread.authorize(credentials)
        print('Authentification is successfull')
        wks = gs.open_by_key(ssheet_id)
        worksheet = wks.get_worksheet(0)
        llists = worksheet.get_all_values()
        print('Data is loaded')
    except gspread.exceptions.SpreadsheetNotFound as e:
        raise e

    if len(llists) == 0:
        raise ValueError('Spreadsheet is empty')
    else:
        headers = llists[0]
    return worksheet, headers, llists[1:]


def calculate_moving_average(calc_window):
    worksheet, headers, llists = get_data(credentials_json, ssheet_id, scope)
    if len(llists) <= calc_window:
        print('Not enough data to calculate moving average with window %s' % (calc_window))
        return
    visitor_number = get_column_values(llists, headers, 'Visitors', int)
    dates = get_column_values(llists, headers, 'Date')
    visitor_number_ma = moving_average(visitor_number, calc_window)
    add_new_column(worksheet, len(llists[0]) + 1, visitor_number_ma, 'Moving Average {}'.format(calc_window), calc_window)


def get_column_values(llists, headers, column_name, value_type=str):
    i = headers.index(column_name)
    list_of_values = [value[i] for value in llists]
    if value_type == int:
        return list(map(lambda x: cast_int(x, 0), list_of_values))
    return list_of_values


def add_new_column(worksheet, number_of_cols, data_to_add, column_name, calc_window):
    worksheet.add_cols(1)
    number_of_rows = len(data_to_add) + 1
    symbol = get_col_name(number_of_cols)
    cell_list = worksheet.range('%s%s' % (symbol, 1)) + \
                worksheet.range('%s%s:%s%s' % (symbol, calc_window + 1 , symbol, number_of_rows + calc_window - 1 ))
    values = [column_name] + [str(v) for v in data_to_add]
    for cell, value in zip(cell_list, values):
        cell.value = value
    worksheet.update_cells(cell_list)
    print('Calculation finished')


def cast_int(x, default_int_value):
    try:
        return int(x)
    except ValueError as e:
        return default_int_value


def moving_average(values, calc_window):
    weights = np.repeat(1.0, calc_window) / calc_window
    ma = np.convolve(values, weights, 'valid')
    ma = [int(round(x)) for x in ma]
    return ma


def get_col_name(col):
    spreadsheet_col = str()
    div = col
    while div:
        (div, mod) = divmod(div-1, 26) # 26 - numbers of letters in alphabet
        spreadsheet_col = chr(mod + 65) + spreadsheet_col
    return spreadsheet_col


if __name__ == '__main__':
   configpath = 'config.txt'
   configParser = configparser.RawConfigParser()
   configParser.read(configpath)
   credentials_json = str(configParser.get('arguments', 'credentials_json'))
   calc_window = int(configParser.get('arguments', 'calc_window'))
   ssheet_id = configParser.get('arguments', 'ssheet_Id')
   scope = configParser.get('arguments', 'scope')

   parser = argparse.ArgumentParser()
   parser.add_argument('-id', type=str, help='Google Spreadsheet id')
   args = vars(parser.parse_args())
   if args['id']:
       ssheet_id = args['id']

   calculate_moving_average(calc_window)
