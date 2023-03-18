## Donate

If this project helped you and you'd like to donate:

[![paypal](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/donate/?hosted_button_id=SX6XF7L3H8GS4)

# ebay-transaction-parser
When Easy Auction Tracker decided to close shop I looked around for an alternative and couldn't find anything that I liked so I decided to build my own. This is a python script that will take an eBay transaction report and import orders, shipping costs, fees, and display it nicely in a master report for tracking purposes. 

![alt text](https://github.com/osirisad/ebay-transaction-parser/blob/master/sample.png?raw=true)

## Installation

Download the excel file template and python script, ensure python script is in the same directory as excel file before running. You should have an import folder and archive folder in the same directory.

Before running, make sure you you install python (I'm using version 3.11.2 - https://www.python.org/ftp/python/3.11.2/python-3.11.2-amd64.exe)

After you install python, install module openxyl by opening command prompt and typing: 
```
pip install openpyxl
```
## Usage

Download transaction reports from [https://www.ebay.com/sh/fin/report](https://www.ebay.com/sh/fin/report) and save them to the import folder.  This script can handle reports from multiple ebay accounts.

Once you're ready to run the report open a command prompt navigate to the folder where the python script is saved then type the following:
python ebay_report.py or double click the run_reports.bat file that's here in the repo.

I've been making a backup of the master file before I update it just in case, you probably sould too! Maybe I will have the script do that in the future, for now make sure you back it up!


