Sptest
======
Sptest lets you to get futures data form [Quandl](http://www.quandl.com/futures) using spreadmatrix.xlsx file. Each data array is cached in cache-data folder.
In output sptest provides calculations of various values, i.e. counts of poitive and negative deals, standard deviation, profit of deals, Calmar, Sterling, Sharpe, Sortino ratios, skewness and kurtosis etc. The script also builds different charts such as total cummulative chart of all deals, chart of drawdowns, VAMI, VaR, Omega and normal distribution charts.

Environment Setup
------
Before running the scripts you need to install some additional python modules:
- [Quandle](http://www.quandl.com/help/python),
- [numpy](http://www.scipy.org/install.html),
- [Mathplotlib](http://matplotlib.org/users/installing.html),
- [openpyxl](http://openpyxl.readthedocs.org/en/2.0.2/),
- [pandas](http://pandas.pydata.org/pandas-docs/stable/install.html),
- [xlsxwriter](http://xlsxwriter.readthedocs.org/en/latest/getting_started.html)

Run the scripts
------
You can run scripts with a periods of years or without it. The common command to run the script is:

    python scriptname.py excelfilename.xlsx startyear endyear
where <i>startyear</i> and <i>endyear</i> are optional.  

There are 4 different scripts for building and calculating various things. 
####sptest_multicharts.py
Is displays total cumulative chart for each row from spreadmatrix.xlsx.
To run the script use following commands:

    python sptest_multicharts.py spreadmatrix.xlsx 
    python sptest_multicharts.py spreadmatrix.xlsx 2000 2005
####sptest_eachRow.py
Is used for calculating coeficients which are multipliers of deals in the main script. You can run this script first and than add one of the 3 coeficients to spreadmatrix.xlsx.
To run the script use following commands:

    python sptest_multicharts.py spreadmatrix.xlsx 100
    python sptest_multicharts.py spreadmatrix.xlsx 1999 2013 100
where <i>100</i> is the divider for coeficients. It can be any natural number. 
The results are stored in <b>coefficients.xlsx</b>
####sptest_report.py
This is the main script and it is used for calculating main values and drawing all charts.
To run the script use following commands:

    python sptest_multicharts.py spreadmatrix.xlsx 100
    python sptest_multicharts.py spreadmatrix.xlsx 1999 2013 100
You can also run it only with one specific contract:

    python sptest_multicharts.py spreadmatrix.xlsx 1999 2013 ICE/B 100
where <i>ICE/B</i> is a contract.
The results are stored in <b>reports.xlsx</b> file.
####sptest_unrealizedProfit.py
Is used for calculating and draw chart of unrealized profit. 
To run the script use following commands:

    python sptest_multicharts.py spreadmatrix.xlsx
    python sptest_multicharts.py spreadmatrix.xlsx 1999 2013
The result is stored in <b>unrealized_profit.xlsx</b>

