# -*- coding: utf-8 -*-
"""
Created on Wed Jul 13 07:52:00 2022
@author: Gilbr@
This file contains all datas for the program.
Variables and imports are made here for avaiding ambiguity.
"""
import os
import oschmod                # For file access mode
import textwrap
import shutil                 # For moving file
from datetime import datetime
import re                     # For regular expression 
import logging

import win10toast             # For notification
import pandas as pd           # For dataframes manupulation

# The base dir where files (inputs and outputs) are saved
basePath = "E:/Gilbert BEMWIZ/py_projects/PayTrans/"

#oschmod.set_mode(basePath, 0o775) # If we'd like to limit access


# Logging for all messages occured
# If we'd like save logs by date, we should create variable for saving current time
#log_temps = f"{datetime.now().strftime('%d_%m_%Y %H_%M_%S')}" 
# Create logger
logger = logging.getLogger('app_logs')
logger.setLevel(logging.DEBUG)
# Create file handler and set level to debug
#logHandler = logging.FileHandler(f"{basePath}/Messages/log_{log_temps}.log") # If by time
logHandler = logging.FileHandler(f"{basePath}/Utils/log_messages.log")
logHandler.setLevel(logging.DEBUG)
# Create formatter
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%d/%m/%Y %H:%M:%S')
# Add formatter to logHandler
logHandler.setFormatter(formatter)
# Add logHandler to logger
logger.addHandler(logHandler)

def removeSpecialChars(value):
    """Use regex for deleting special chars"""
    return re.sub(r'[^a-zA-Z0-9-/-_]', '', value)

error_msg = ""

# Take all bank account lists in the basPath/Input_ExcelFiles/Bank_Account_List/
bankAccountList = [
    f for f in os.listdir(basePath+ "Input_ExcelFiles/Bank_Account_List/") \
        if os.path.isfile(os.path.join(basePath+ "Input_ExcelFiles/Bank_Account_List/", f))
        ]
# Take the Excel files
if bankAccountList:
    bankAccountList = pd.read_excel(
        basePath + "Input_ExcelFiles/Bank_Account_List/" + bankAccountList[0],
        usecols=['No.', 'Name', 'Bank Account No.']
        ).fillna(value='')
else:
    error_msg += "\nAny Bank Account List"

# Same thing for vendor bank account lists
vendorBankAccountList = [f for f in os.listdir(basePath + "Input_ExcelFiles/Vendor_Bank_Account_List/") if os.path.isfile(os.path.join(basePath+ "Input_ExcelFiles/Vendor_Bank_Account_List/", f))]
if vendorBankAccountList:
    vendorBankAccountList = pd.read_excel(
        basePath + "Input_ExcelFiles/Vendor_Bank_Account_List/" + vendorBankAccountList[0],
        usecols=['Vendor No.', 'Vendor Name', 'Bank Account No.', 'SWIFT Code', 'IBAN', 'Currency Code']
        ).fillna(value='')
else:
    error_msg += "\nAny Vendor Bank Account List"
   
vendorList = [f for f in os.listdir(basePath + "Input_ExcelFiles/Vendor_List/") if os.path.isfile(os.path.join(basePath+ "Input_ExcelFiles/Vendor_List", f))]
if vendorList:
    vendorList = pd.read_excel(
        basePath + "Input_ExcelFiles/Vendor_List/" + vendorList[0],
        usecols=['No.', 'Name', 'Address']
        ).fillna(value='')
else:
    error_msg += "\nAny Vendor List"

# We take all payment proposals
# For TMB
allPayementProposal = [f for f in os.listdir(basePath + "Input_ExcelFiles/Payment_Proposal/TMB/") if os.path.isfile(os.path.join(basePath+ "Input_ExcelFiles/Payment_Proposal/TMB/", f))]
allPayementProposal = [n for n in allPayementProposal if "~" not in n]

payementProposalTmbAutresCDF = [f for f in allPayementProposal if "CDF" in f and "TMB" not in f]
payementProposalTmbAutresUSD = [f for f in allPayementProposal if "USD" in f and "TMB" not in f]
payementProposalTmbTmbCDF = [f for f in allPayementProposal if "CDF" in f and "TMB" in f]
payementProposalTmbTmbUSD = [f for f in allPayementProposal if "USD" in f and "TMB" in f]

# For Ecobank
allPayementProposal = [f for f in os.listdir(basePath + "Input_ExcelFiles/Payment_Proposal/Ecobank/") if os.path.isfile(os.path.join(basePath+ "Input_ExcelFiles/Payment_Proposal/Ecobank/", f))]
# We exclude temporal files
allPayementProposal = [n for n in allPayementProposal if "~" not in n]
# We devide our payment proposals as we're gonna make outputs
payementProposalEcoAutresCDF = [f for f in allPayementProposal if "CDF" in f and "Ecobank" not in f]
payementProposalEcoAutresUSD = [f for f in allPayementProposal if "USD" in f and "Ecobank" not in f]
payementProposalEcoEcoCDF = [f for f in allPayementProposal if "CDF" in f and "Ecobank" in f]
payementProposalEcoEcoUSD = [f for f in allPayementProposal if "USD" in f and "Ecobank" in f]

# For debug
print(error_msg)


# In logs file, we inform user all errors occured 
if error_msg:
    logger.error(error_msg)