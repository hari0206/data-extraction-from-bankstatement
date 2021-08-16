from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
from dateutil.parser import parse
from datetime import datetime
import os
import xlrd
import json
import pandas as pd



def get_bank_values(bank_name):
    with open("config.json", "r") as f:
        content = json.loads(f.read())
    all_banks = content["banks"]
    for bank in all_banks:
        if bank.get("name") == bank_name.lower():
            return bank["header"]


def extract_data_from_statement(file_path):
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_index(0)

    data = []
    for row in range(worksheet.nrows):
        row_data = []
        for col in range(worksheet.ncols):
            value = worksheet.cell_value(row, col)
            if len(str(value).strip()) > 0:
                row_data.append(value)
        data.append(row_data)

    return data


def extract_raw_data_from_statement(file_path):
    workbook = xlrd.open_workbook(file_path)
    worksheet = workbook.sheet_by_index(0)

    data = []
    for row in range(worksheet.nrows):
        row_data = []
        for col in range(worksheet.ncols):
            value = worksheet.cell_value(row, col)
            row_data.append(value)
        data.append(row_data)
    return data


def match_headers(headers, data):
    header_list_bank = list(headers.values())
    header_list = [
        header for header in header_list_bank if len(header.strip()) > 0]
    for index, values in enumerate(data):
        values = [ele.strip() for ele in values if type(ele) == str]
        if values == header_list:
            return index


def find_end_index(data):

    for index, values in enumerate(data):
        flag = False
        for element in values:
            if len(str(element).strip()) > 0:
                flag = True
        if not flag:
            return index
    return len(data)


def match_keys(data, headers):
    final_headers = []
    data_headers = data[0]
    for dh in data_headers:
        if len(dh.strip()) > 0:
            for h in headers:
                if headers.get(h) == dh:
                    final_headers.append(h)

    for h in headers:
        if not h in final_headers:
            final_headers.append(h)

    data[0] = final_headers
    return data


def is_date(string, fuzzy=False):
    try:
        parse(string, fuzzy=fuzzy)
        return True
    except ValueError:
        return False


def excel_date_format(json_data):
    excel_date = [date["Date"] for date in json_data["data"]]
    format_date = is_date(str(excel_date[0]))
    formated_date = []
    if format_date == False:
        for x in excel_date:
            if x:
                dt = datetime.fromordinal(
                    datetime(1900, 1, 1).toordinal() + int(x) - 2)
                formated_date.append(dt)
        return formated_date
    else:
        return excel_date


def extract_data(bank_name, statement_path):
    headers = get_bank_values(bank_name)
    data = extract_data_from_statement(statement_path)
    start_index = match_headers(headers, data)
    end_index = find_end_index(data[start_index:])
    print(start_index,end_index)
    raw_data = extract_raw_data_from_statement(statement_path)
    data = raw_data[start_index:start_index+end_index]
    for x in data:
        for i, y in enumerate(x):
            if y == "":
                x[i] = None

    df = pd.DataFrame(data, index=None)
    df = df.dropna(axis=1, how='all')
    data = df.values.tolist()
    data = match_keys(data, headers)
    df = pd.DataFrame(data)
    df.columns = data[0]
    df = df[1:]
    print(df)
    json_data = df.to_json(orient='table', index=False)
    json_data = json.loads(json_data)
    for x in json_data["data"]:
        if (x['Withdrawal Amt.'] == None or not str(x['Withdrawal Amt.']).strip()) == (x['Deposit Amt.'] == None or not str(x['Deposit Amt.']).strip()):
            json_data["data"].remove(x)
        elif x['Withdrawal Amt.'] == None and x['Deposit Amt.'] == None:
            json_data["data"].remove(x)
    formated_date = excel_date_format(json_data)
    # print(is_date(str(json_data["data"][0]["Date"])))
    if is_date(str(json_data["data"][0]["Date"])) == False:
        for date, transaction_date in zip(formated_date, json_data["data"]):
            transaction_date.update({"Date": date.date(),
                                     "Value Dt": date.date()})
    return json_data['data']


statement_path = r"C:\Users\lenovo\OneDrive - BOT Mantra\Desktop\data-extraction-from-bankstatement\statements\sbi_statement.xlsx"
bank_name = "SBI"
# statement_path = r"C:\Users\lenovo\OneDrive - BOT Mantra\Desktop\Bank Statements\hdfc_statement.xls"
# bank_name = "hdfc"
# statement_path = r"C:\Users\lenovo\OneDrive - BOT Mantra\Desktop\Bank Statements\Indian Bank.xls"
# bank_name = "indian"
# statement_path = r"C:\Users\lenovo\OneDrive - BOT Mantra\Desktop\Bank Statements\CWS-Bank Of Baroda.xlsx"
# bank_name = "bankofbaroda"
# statement_path = r"C:\Users\lenovo\OneDrive - BOT Mantra\Desktop\Bank Statements\ICICIOpTransactionHistory13-08-2021.xls"
# bank_name = "icici"
# statement_path = r"C:\Users\lenovo\OneDrive - BOT Mantra\Desktop\Bank Statements\CWS-Canara Bank.xlsx"
# bank_name = "canara"
# statement_path = r"C:\Users\lenovo\OneDrive - BOT Mantra\Desktop\Bank Statements\dbs_statement.xls"
# bank_name = "dbs"
# statement_path = r"C:\Users\lenovo\OneDrive - BOT Mantra\Desktop\Bank Statements\CWS-AXIS.xlsx"
# bank_name = "axis"


print(extract_data(bank_name,statement_path))