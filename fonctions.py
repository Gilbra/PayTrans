# -*- coding: utf-8 -*-
"""
Created on Wed Jul 13 10:02:54 2022

@author: Gilbr@
"""
from datas import *

def translateTmbAutre(payementProposalTmbAutres, devise='CDF'):
    """
    Function for TMB AUTRE banque. payementProposalTmbAutres is the 
    list of payment proposals and the devise is whether CDF or other.
    The function take the list, convert each list element ie a payment proposal,
    loop all payment proposal elements takes informations in others files 
    (vendor list, bank account list..) and writes an text file.
    """
    error_msg = ""
    if len(payementProposalTmbAutres) < 1:
        logger.info("Any payment proposal for TMB-AUTRES ")
    for file in payementProposalTmbAutres:
        try:
            # We take only necessaries culumns
            paymentProposal = pd.read_excel(
                basePath + "Input_ExcelFiles/Payment_Proposal/TMB/" + file,
                usecols=[
                    'Account No.', 'Amount',
                    'Amount LCY DRC', 'IBAN',
                    'Applies-to Ext. Doc. No.','Applies-to Doc. No.'
                    ],
                #engine='pyxlsb'
                ).fillna(value='')
            
            text = "TMB_AUTRES/AUTRES/"
            if devise == 'CDF':
                text = "TMB_AUTRES/CDF/"
            # The text file for output
            fileName = basePath + "Output_TextFiles/" + text + datetime.now().strftime("%Y%m%d%H%M%S") + "_TMB_AUTRE_"+ devise +".txt"
            with open(fileName, "w") as sortie:
                # Look for the debit account
                if devise == 'CDF':
                    debitAccount = bankAccountList.loc[bankAccountList['Name'] == "TMB CDF LUBUMBASHI"]["Bank Account No."]
                else:
                    debitAccount = bankAccountList.loc[bankAccountList['Name'] == "TMB USD LUBUMBASHI"]["Bank Account No."]
                # If any debit account is found
                if not len(debitAccount):
                    # Raise an error
                    error_msg += 'Pas de compte à debiter pour TMB USD LUBUMBASHI dans Bank Account List\n'
                    logger.critical("Pas de compte à debiter pour TMB USD LUBUMBASHI dans Bank Account List")
                    continue
                debitAccount = debitAccount.iloc[0]
                debitAccount = debitAccount[:5]+"-"+debitAccount[5:10]+"-"+debitAccount[10:21]+"-"+debitAccount[21:]
                
                for index, row in paymentProposal.iterrows():
                    if devise == 'CDF':
                        montant = row["Amount LCY DRC"]
                    else:
                        montant = row["Amount"]
                    ligne = f"A;0027349;{debitAccount};{montant:.2f};"
                    
                    vendor = vendorBankAccountList.loc[(vendorBankAccountList['IBAN'] == row['IBAN']) & (vendorBankAccountList['Vendor No.'] == row['Account No.'])]
                    # For test we wrote this line above when we didn't found datas
                    #vendor = vendorBankAccountList.loc[vendorBankAccountList['Vendor No.'] == row['Account No.']]
                    if not len(vendor):
                        # For test
                        vendor = vendorBankAccountList
                        # Raise an error
                        error_msg += f"{row['Account No.']} Pas de données correspondantes dans Vendor Bank Account List\n"
                        logger.warning(f"{row['Account No.']} Pas de données correspondantes dans Vendor Bank Account List. ligne {index + 1} {file}")
                        continue
                    vendor = vendor.iloc[0]
                    currency_code = vendor['Currency Code']
                    if not currency_code:
                        currency_code = "CDF"
                    ligne += f"{currency_code};"
                    ligne += f"{textwrap.shorten(vendor['Vendor Name'], width=35, placeholder='..')};"
                    address = vendorList.loc[vendorList['Name'] == vendor['Vendor Name']]["Address"]
                    if not len(address):
                        # Raise an error
                        error_msg += 'Pas d\'adresse correspondante dans Vendor List\n'
                        logger.warning(f"Pas d\'adresse correspondante dans Vendor List. ligne {index + 1} {file}")
                        continue
                    address = address.iloc[0]
                    ligne += f"{address};{address};"
                    creditAccount = row['IBAN']
                    if not isinstance(creditAccount, str):
                        creditAccount = f"{creditAccount:023}"
                    ligne += creditAccount + ";"
                    ligne += f"{vendor['SWIFT Code']};"
                    desc = f"{row['Applies-to Ext. Doc. No.']}{row['Applies-to Doc. No.']}"
                    ligne += removeSpecialChars(desc) + ";" + removeSpecialChars(desc)
                    ligne += '\n'
                    sortie.write(ligne)
                    
                sortie.close()
                # File mode access : we make the output only readable
                oschmod.set_mode(fileName, 0o555)
            shutil.move(basePath + "Input_ExcelFiles/Payment_Proposal/" + file, basePath + "Archives/" + text + file)
            logger.info(f"{file} >> {fileName} OK")
        except PermissionError:
            error_msg += f"{file} est ouvert dans un autre programme, impossible de l archiver"
            logger.error(f"{file} est ouvert dans un autre programme, impossible de l archiver")
        except:
            logger.error(f"{file} est ouvert dans un autre programme, impossible de l archiver")
        finally:
            pass
    print(error_msg)
    return

def translateTmbTmb(payementProposalTmbTmb, devise='CDF'):
    """
    Function for TMB TMB. payementProposalTmbAutres is the 
    list of payment proposals and the devise is whether CDF or other.
    The function take the list, convert each list element ie a payment proposal,
    loop all payment proposal elements takes informations in others files 
    (vendor list, bank account list..) and writes an text file.
    """
    error_msg = ""
    if len(payementProposalTmbTmb) < 1:
        logger.inf("Any payment proposal for TMB-TMB ")
    try:
        for file in payementProposalTmbTmb:
            paymentProposal = pd.read_excel(
                basePath + "Input_ExcelFiles/Payment_Proposal/TMB/" + file,
                usecols=[
                    'Account No.', 'Amount',
                    'Amount LCY DRC', 'IBAN',
                    'Applies-to Ext. Doc. No.','Applies-to Doc. No.'
                    ],
                #engine='pyxlsb'
                ).fillna(value='')
            
            text = "TMB_TMB/AUTRES/"
            if devise == 'CDF':
                text = "TMB_TMB/CDF/"
            
            fileName = basePath + "Output_TextFiles/" + text + datetime.now().strftime("%Y%m%d%H%M%S") + "_TMB_TMB_"+ devise +".txt"
            with open(fileName, "w") as sortie:
                if devise == 'CDF':
                    debitAccount = bankAccountList.loc[bankAccountList['Name'] == "TMB CDF LUBUMBASHI"]["Bank Account No."]
                else:
                    debitAccount = bankAccountList.loc[bankAccountList['Name'] == "TMB USD LUBUMBASHI"]["Bank Account No."]
                if not len(debitAccount):
                    # Raise an error
                    error_msg += 'Pas de compte à debiter pour TMB USD LUBUMBASHI dans Bank Account List\n'
                    logger.critical("Pas de compte à debiter pour TMB USD LUBUMBASHI dans Bank Account List")
                    continue
                debitAccount = debitAccount.iloc[0]
                if not isinstance(debitAccount, str):
                    debitAccount = f"{debitAccount:023}"
                debitAccount = debitAccount[:4]+"-"+debitAccount[4:11]+"-"+debitAccount[11:13]+"-"+debitAccount[13:]
                if devise == 'CDF':
                    totalMontant = paymentProposal["Amount LCY DRC"].sum()
                else:
                    totalMontant = paymentProposal["Amount"].sum()
                    
                ligne = f"0027349;{debitAccount};{totalMontant:.2f};{devise}\n"
                sortie.write(ligne)
                
                for index, row in paymentProposal.iterrows():
                    if devise == 'CDF':
                        montant = row["Amount LCY DRC"]
                    else:
                        montant = row["Amount"]
                    creditAccount = row['IBAN']
                    if not isinstance(creditAccount, str):
                        creditAccount = f"{creditAccount:023}"
                    creditAccount = creditAccount[:4]+"-"+creditAccount[4:9]+"-"+creditAccount[11:20]+"-"+creditAccount[20:]
                    ligne = f"A;{montant:.2f};{creditAccount};"
                    desc = f"{row['Applies-to Ext. Doc. No.']}{row['Applies-to Doc. No.']}"
                    ligne += removeSpecialChars(desc) + ";" + removeSpecialChars(desc)
                    
                    ligne += ';SAL\n'
                    sortie.write(ligne)
                    
                sortie.close()
                oschmod.set_mode(fileName, 0o555)
            shutil.move(basePath + "Input_ExcelFiles/Payment_Proposal/" + file, basePath + "Archives/" + text + file)
            logger.info(f"{file} >> {fileName} OK")
    except PermissionError:
        error_msg += f"{file} est ouvert dans un autre programme, impossible de l archiver"
        logger.error(f"{file} est ouvert dans un autre programme, impossible de l archiver")
    except:
        logger.error(f"{file} est ouvert dans un autre programme, impossible de l archiver")
    finally:
        pass    
    print(error_msg)
    return

def translateEcobank(payementProposalEcobankAutre, devise='CDF', vers_eco=False):
    """
    Function for ECOBANK AUTRE or ECOBANK ECOBANK. payementProposalTmbAutres is the 
    list of payment proposals and the devise is whether CDF or other. And vers_eco tells us
    if transaction is between ECOBANK accounts (True) or between an ECOBANK acount and an account of 
    other bank (False).
    The function take the list, convert each list element ie a payment proposal,
    loop all payment proposal elements takes informations in others files 
    (vendor list, bank account list..) and writes an text file.
    """
    error_msg = ""
    if len(payementProposalEcobankAutre) < 1:
        if vers_eco:
            logger.inf("Any payment proposal for ECOBANK-ECOBANK ")
        else:
            logger.inf("Any payment proposal for ECOBANK-AUTRES ")
    for file in payementProposalEcobankAutre:
        try:
            paymentProposal = pd.read_excel(
                basePath + "Input_ExcelFiles/Payment_Proposal/Ecobank/" + file,
                usecols=[
                    'Account No.', 'Amount',
                    'Amount LCY DRC', 'IBAN',
                    'Applies-to Ext. Doc. No.','Applies-to Doc. No.'
                    ],
                #engine='pyxlsb'
                ).fillna(value='')
            
            text = "ECOBANK_AUTRES/"
            if vers_eco:
                text = "ECOBANK_ECOBANK/"
            if devise == 'CDF':
                text += "CDF/"
            else:
                text += "AUTRES/"
            
            fileName = basePath + "Output_TextFiles/" + text + datetime.now().strftime("%Y%m%d%H%M%S") + "_ECO_AUTRE_"+ devise +".txt"
            # If it's ECOBANK-ECOBANK
            if vers_eco:
                fileName = basePath + "Output_TextFiles/" + text + datetime.now().strftime("%Y%m%d%H%M%S") + "_ECO_ECO_"+ devise +".txt"
            with open(fileName, "w") as sortie:
                temps = datetime.now().strftime('%d%m%Y')
                BatchReference = f"ECO-AUTRES {devise}"
                if vers_eco:
                    BatchReference = f"ECOBANK {devise}"
                ligne = f"H,ECOCORP,Paiement Fournisseurs,,,,{temps},{BatchReference},,,\n"
                sortie.write(ligne)
                if devise == 'CDF':
                    debitAccount = bankAccountList.loc[bankAccountList['Name'] == "ECOBANK CDF KINSHASA"]["Bank Account No."]
                else:
                    debitAccount = bankAccountList.loc[bankAccountList['Name'] == "ECOBANK USD KINSHASA"]["Bank Account No."]
                if not len(debitAccount):
                    # Raise an error
                    error_msg += f"Pas de compte à debiter pour ECOBANK {devise} KINSHASA dans Bank Account List\n"
                    logger.critical("Pas de compte à debiter pour ECOBANK {devise} KINSHASA dans Bank Account List")
                    continue
                debitAccount = debitAccount.iloc[0]
                #debitAccount = debitAccount[:4]+"-"+debitAccount[4:11]+"-"+debitAccount[11:13]+"-"+debitAccount[13:]
                if devise == 'CDF':
                    totalMontant = paymentProposal["Amount LCY DRC"].sum()
                else:
                    totalMontant = paymentProposal["Amount"].sum()
                    
                
                for index, row in paymentProposal.iterrows():
                    ligne = f"D,{index + 1},,,"
                    #vendor = vendorBankAccountList.loc[(vendorBankAccountList['IBAN'] == row['IBAN']) & (vendorBankAccountList['Vendor No.'] == row['Account No.'])]
                    vendor = vendorBankAccountList.loc[vendorBankAccountList['Vendor No.'] == row['Account No.']]
                    if not len(vendor):
                        # For test
                        vendor = vendorBankAccountList
                        # Raise an error
                        error_msg += f"{row['Account No.']} Pas de données correspondantes dans Vendor Bank Account List\n"
                        logger.warning(f"{row['Account No.']} Pas de données correspondantes dans Vendor Bank Account List")
                        continue
                    vendor = vendor.iloc[0]
                    BenBankID = "Bank Sort Code"
                    if vers_eco:
                        BenBankID = "CD026001"
                    ligne += f"{textwrap.shorten(vendor['Vendor Name'], width=35, placeholder='..')},,,SYSTEM,{BenBankID},,,,ECOBANK,DRC,{row['IBAN']},"
                    currency_code = vendor['Currency Code']
                    if not currency_code:
                        currency_code = "CDF"
                    ligne += f"{currency_code},,,,,"
                    address = vendorList.loc[vendorList['Name'] == vendor['Vendor Name']]["Address"]
                    if not len(address):
                        # Raise an error
                        error_msg += 'Pas d\'adresse correspondante dans Vendor List\n'
                        logger.warning('Pas d\'adresse correspondante dans Vendor List')
                        continue
                    address = address.iloc[0]
                    if devise == 'CDF':
                        montant = row["Amount LCY DRC"]
                    else:
                        montant = row["Amount"]
                    ligne += f"{address},,,,,,,,,,,,,,,,,,,,,,,{debitAccount},{devise},,,,{montant:.2f},,{temps},{temps},,,,,,,,,,"
                    
                    desc = f"{row['Applies-to Ext. Doc. No.']}{row['Applies-to Doc. No.']}"
                    ligne += removeSpecialChars(desc) + ",," + removeSpecialChars(desc)
                    
                    ligne += f",,,,,,,,,,,,,,,,,,,,,\n"
                    
                    sortie.write(ligne)
                    
                ligne += f"T,{len(paymentProposal)},{totalMontant:.2f},\n"
                sortie.write(ligne)
                
                sortie.close()
                oschmod.set_mode(fileName, 0o555)
                
            shutil.move(basePath + "Input_ExcelFiles/Payment_Proposal/Ecobank/" + file, basePath + "Archives/" + text + file)
            logger.info(f"{file} >> {fileName} OK")
            return
        except PermissionError:
            error_msg += f"{file} est ouvert dans un autre programme, impossible de l'archiver"
            logger.error(f"{file} est ouvert dans un autre programme, impossible de l archiver")
            pass
        except ConnectionRefusedError as c:
            logger.error(f"{c}")
            pass
        except Exception as e:
            logger.error(f"{e}")
            pass
        finally:
            pass

def main():
    """
    All functions are called here. The main of the program.
    """
    # TMB vers Autres banques
    translateTmbAutre(payementProposalTmbAutresCDF, devise='CDF')
    translateTmbAutre(payementProposalTmbAutresUSD, devise='USD')
    
    # TMB vers TMB
    translateTmbTmb(payementProposalTmbTmbCDF, devise='CDF')
    translateTmbTmb(payementProposalTmbTmbUSD, devise='USD')
    #"""
    # ECOBANK vers autres banques
    translateEcobank(payementProposalEcoAutresCDF, devise='CDF')
    translateEcobank(payementProposalEcoAutresUSD, devise='USD')
    
    # ECOBANK vers ECOBANK
    translateEcobank(payementProposalEcoEcoCDF, devise='CDF', vers_eco=True)
    translateEcobank(payementProposalEcoEcoUSD, devise='USD', vers_eco=True)
