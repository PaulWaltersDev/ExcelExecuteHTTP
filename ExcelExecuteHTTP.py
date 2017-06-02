import os.path
import sys
from datetime import datetime as dt

import requests
from openpyxl import load_workbook

# ExcelExecuteHTTP
# Paul Maxwell-Walters, May 2017

__author__ = 'Paul Maxwell-Walters'
DEFAULT_WORKBOOK = 'Excel-Execute-HTTP-Test.xlsx'
DEFAULT_RESULTS_FILE = 'results.txt'


def load_url_data(excel_file, sheetname="Sheet1"):
    """
    Parse Excel file for URL paths and payloads to test
    :param excel_file: path to Excel workbook to load
    :param sheetname: name of workbook sheet to read, default Sheet1
    :return: list of tuples containing data for testing, e.g.
        (URL, HTTP Method, param1, value1, param2, value2)
    """
    if not os.path.exists(excel_file):
        raise FileNotFoundError(f'File {excel_file} does not exist.')
    try:
        workbook_data = load_workbook(filename=excel_file, data_only=True)
        return tuple(workbook_data[sheetname].rows)
    except IOError as exc:
        print(repr(exc))
        print("Error loading or parsing excel file, or Excel file is invalid.")
        exit(-1)


def execute_requests(url_data, output_file=None):
    """
    Make HTTP requests to URLs with the methods and params provided in workbook data
    :param url_data: data loaded from workbook for testing, see docstring for load_url_data
    :param output_file: string location of output file for results, None for stdout-only
    """
    print_and_save(f'START SESSION {str(dt.now())}', file=output_file)
    for target in url_data:
        url = target[0].value
        method = str.upper(target[1].value)
        params = dict()

        for column in range(2, len(target), 2):
            if target[column].value and target[column + 1].value:
                params[target[column].value] = target[column + 1].value

        print_and_save("--------------------------------------------------------------", file=output_file)
        print_and_save(f'Executing Request for URL: {url} -------- Payload {params}', file=output_file)
        if method == 'GET':
            save_response(requests.get(url=url, params=params), output=output_file)
        elif method == 'POST':
            save_response(requests.post(url=url, params=params), output=output_file)
        else:
            print_and_save(f'{method} not allowed, only GET and POST can be used.\n\n', file=output_file)
            continue
    print_and_save(f'END SESSION {str(dt.now())}\n\n\n\n', file=output_file)


def save_response(response, output):
    """
    Print the response data we're interested in.
    :param response: requests.Response object
    :param output: desired output target in addition to stdout
    """
    result = f'URL = {response.url}\n' \
             f'Status Code = {response.status_code}\n' \
             f'Headers = {response.headers}\n' \
             f'Response = {response.text}\n' \
             f'History = {response.history}\n' \
             f'Time Elapsed = {response.elapsed.total_seconds()}s'
    print_and_save(result, file=output)
    for cookie in response.cookies:
        print(f'Cookie = {cookie}', file=output)
    print_and_save('\n\n', file=output)


def print_and_save(line, file):
    """
    Print lines to stdout as well as output file.
    :param line: line to print
    :param file: output file to write to, skip if None
    """
    print(line)
    if file:
        with open(file, 'a') as outfile:
            outfile.write(f'{line}\n')


def read_workbook_and_execute_requests(workbook, output_file):
    """
    Read test data from Excel file and execute the calls.
    :param workbook: path to Excel file
    :param output_file: path to results file, None if not required.
    """
    url_submit_data = load_url_data(workbook)
    execute_requests(url_data=url_submit_data, output_file=output_file)


if __name__ == '__main__':
    """
    Script execution. First argument provided should be the workbook containing test data.
    Use default workbook if no args are provided.
    Writes results to DEFAULT_RESULTS_FILE and stdout by default. Set output_file=None if you
    only want to print the output.
    """
    try:
        wb = sys.argv[1]
    except IndexError as err:
        print(f'No arguments provided, reading data from {DEFAULT_WORKBOOK}...')
        wb = DEFAULT_WORKBOOK
    read_workbook_and_execute_requests(workbook=wb, output_file=DEFAULT_RESULTS_FILE)
