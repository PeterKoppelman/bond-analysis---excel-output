import datetime

# The auth_key is the quandl key that is used to get data. You will need to
# sign up at quandl.com. The libraries that you will be using for this are free. 
# Check on the free filter and the real estate asset class to get the first three libaraies. 
# You need the following data:
# 	1) Freddie Mac (30-Year Fixed Rate Mortgage Average in the United States). 
#	   Click on Libraries: Python
# 	2) Wells Fargo Home Mortgage Loans:
#		a) Home Mortgage Loans: Purchase Rate, Conforming Loan, 30-Year Fixed Rate, Interest Rate.
#		Click on libraries: Python
# 	3) Wells Fargo Home Mortgage Loans:
#		a) Home Mortgage Loans: Purchase Rate, Jumbo Loan (Amounts that exceed conforming loan limits), 30-Year Fixed Rate, Interest Rate
#		Click on libraries: Python
# For the last two libraries, we're going to get data from the Treasury Department. The asset class for this is
# Interest Rates and Futures. We want Treasury Yield Curve Rates. The library is Python. We'll get daily data
# and the delta (difference) between rates from one day to another. We get the differernce
# by using transform = diff when we grab the data.

# Put auth key here
auth_key = ""

title = ["1 Month Delta", "3 Month Delta", "6 Month Delta", '1 Year Delta',
         '2 Year Delta', '3 Year Delta', '5 Year Delta', '7 Year Delta',
         '10 Year Delta', '20 Year Delta', '30 Year Delta']

chart_cell = ["A1", "K1", "U1", "A19", "K19", "U19", "A37",
              "K37", "U37", "A56", "K56"]

elapsed_seconds = int(datetime.datetime.now().strftime('%H')) * 60 * 60 + \
	int(datetime.datetime.now().strftime('%M')) * 60 + \
	int(datetime.datetime.now().strftime('%S'))

timestamp = datetime.datetime.now().strftime('%Y-%m-%d')
file_name = 'historical delta models {} {}.xlsx'.format(timestamp, elapsed_seconds)
