# ExcelExecuteHTTP

A Quick and Dirty Python 3.5 script to automate the execution of HTTP and HTTPS GET and POST requests with a list of URLs and parameter:value combinations taken from an Excel spreadsheet.

Part of a community contribution submission to COMP6443 Web Security and Testing, Semester 1 2017, UNSW Australia.

This allows the user to use the power of Excel functions to generate any sort of text and numerical payloads using a simple Excel 2010/2013 file format.

##Instructions

 A) Insure that the latest version of Python 3 is installed on your computer, along with the pip install application.

 B) Git-clone the repository. Alternatively download and extract the project files directly from [here](https://github.com/PaulWaltersDev/ExcelExecuteHTTP/archive/master.zip).

 C) Install the two Python libraries required for this script - [Requests](http://docs.python-requests.org/en/master/) and [OpenPyXL](https://openpyxl.readthedocs.io/en/default/). 
 
 D) Create an Excel file with a sheet called "Sheet1" where the columns contain the content and order below -
 
 Column
 1) URL
 2) GET or POST
 3) Parameter
 4) Value
 5),6)... next parameter, value... repeat as appropriate
 
 -- See the included specimen Excel 2013 file Excel-Execute-HTTP-Test.xlsx 

(Note that currently the first row is included in the list of rows to be executed. You must therefore leave out a header row.
This will be amended to allow a header row in the next release.

E) From command line or bash shell, navigate to the extracted files downloaded in step B) and execute -

python (or python3) ExcelExecuteHTTP.py "path to Excel file"

![Python CMD command](https://cloud.githubusercontent.com/assets/10481652/26726310/ce5be39a-47e5-11e7-8243-8d5797239ad7.JPG)

You will see each request executed and the results of each, listed by -

* Request URL
* Response Status Code
* Response Headers
* Response History
* Response Cookies (if any)

![Request and Reasponses](https://cloud.githubusercontent.com/assets/10481652/26726312/ce686bec-47e5-11e7-98fb-41631d7a6d5d.JPG)

The next versions will allow the program output to be saved as a text, CSV or JSON file. For the moment UNIX and Linux users can send to a file
via the suffix " > " output file name and path".

This is quite new and there may be defects I haven't picked up. The functionality is still fairly limited. If you have any defects or questions,
please add to the "Issues" list or contact me on paulwalters2002@yahoo.co.uk.

Paul Maxwell-Walters, [Twitter](https://twitter.com/TestingRants)




