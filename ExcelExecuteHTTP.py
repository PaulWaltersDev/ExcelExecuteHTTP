from openpyxl import load_workbook
import sys
import requests
import os.path

# ExcelExecuteHTTP
# Paul Maxwell-Walters, May 2017

__author__ = 'Paul Maxwell-Walters'
excel_file = sys.argv[1]
work_book = None
url_params_data = None

if os.path.exists(excel_file):

    try:
        work_book = load_workbook(filename=excel_file, data_only=True)
    except IOError:
        print("Error Loading or Parsing Excel file or Excel file invalid")
        exit(-1)

    sheet_ranges = work_book["Sheet1"]  # Currently Sheet1 is hardcoded. Will be amended such that future releases will
                                        # allow sheet name to be added as a parameter

    url_submit_data = tuple(sheet_ranges.rows)

else:
    raise FileNotFoundError("File does not exist or is not an Excel file")

for url_submit in url_submit_data:
    response = None
    url_string = url_submit[0].value
    url_get_or_post = str.upper(url_submit[1].value)
    payload = {}

    print("--------------------------------------------------------------")
    print("Executing Request for URL " + url_string)

    for i in range(2, len(url_submit), 2):  # Adds Parameter Name and value to a Payload Dictionary, missing out empties
        if url_submit[i].value != None and url_submit[i+1].value != None:
            payload[url_submit[i].value] = url_submit[i+1].value

    print("URL " + url_string + " ------------- Payload = " + str(payload))

    if url_get_or_post == "GET":
        response = requests.get(url_string, params=payload)
    elif url_get_or_post == "POST":
        response = requests.post(url_string, params=payload)
    else:
        print("Invalid Request Type: Not GET or POST")
        print("\n" * 2)
        continue

    print("URL = " + str(response.url))     # Prints to output. Will be amended to allow printing back into Excel or into a file ina future release
    print("Status Code = " + str(response.status_code))
    print("Headers = " + str(response.headers))
    print("Response = " + str(response.text))
    print("History = " + str(response.history))
    for cookie in response.cookies:
        print("Cookie = " + cookie)

    print("\n" * 2)





