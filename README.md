# bond-analysis---excel-output
This system gets financial data from quandl,
crunches some numbers and outputs data to an Excel file.

A pip list command shows that I have the following packages and versions on my machine:

Package         Version

certifi         2019.11.28

chardet         3.0.4

cycler          0.10.0

et-xmlfile      1.0.1

idna            2.9

inflection      0.3.1

jdcal           1.4.1

kiwisolver      1.1.0

matplotlib      3.2.1

more-itertools  8.2.0

numpy           1.18.1

openpyxl        3.0.3

pandas          1.0.1

pip             20.0.2

pyparsing       2.4.6

python-dateutil 2.8.1

pytz            2019.3

pywin32         227

Quandl          3.5.0

requests        2.23.0

scipy           1.4.1

seaborn         0.10.0

setuptools      41.2.0

six             1.14.0

urllib3         1.25.8

The data that is inputted from quandl is the following:
  1) Freddie Mac (30-Year Fixed Rate Mortgage Average in the United States). 
  2) Wells Fargo Home Mortgage Loans:
    a) Home Mortgage Loans: Purchase Rate, Conforming Loan, 30-Year Fixed Rate, Interest Rate.
    b) Home Mortgage Loans: Purchase Rate, Jumbo Loan (Amounts that exceed conforming loan limits),
    30-Year Fixed Rate, Interest Rate
  3) Treasury Department Interest Rates and Futures. We want Treasury Yield Curve Rates.
  
You'll need a quandl authentication key for this. Instructions for this are in the auth_token.py file.

Once we have the data, numbers are crunched and output to Excel using the Pandas library ExcelWriter.

Once they're in Excel we use the openpyxl library to create the graphs. The mortgage graphs show the 
difference between jumbo 30 year fixed rate mortgages for purchases and comforming 30 year fixed rate mortgages
for purchases. It's true that jumbo interest rates are lower than conforming interest rates.

The yield curves are a little different. While the mortgage rates graphs are time series graphs, each line in the 
yield curve graph is a time series graph unto itself. There is one yield curve graph for the last day of the most
recent 8 quarters.

The Excel workbook is closed and given a filename with a date/time stamp in it. By using a time stamp,
which is the elapsed seconds since midnight, you can run this several times a day in case there's a data issue.
None of the previous Excel files for the day will be overwritten, which will allow you to compare the excel files
to see if the data issue has been fixed.

The last thing that the system does is email the Excel file that was just created to a distribution list.
Thanks to Michael Shore for writing this code. The email_reference.py file has a distribution list in it.
It's a list, just enter the email addresses of the people that you want to email the Excel workbook to.
There's an emailfrom email address and a password. The last thing that __main__.py does is call the
program email_the_data.py which gets the email addresses from the distribution list and sends emails
to everyone on the list with the most recent Excel workbook as an attachment.
