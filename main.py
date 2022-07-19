# -*- coding: utf-8 -*-
"""
Created on Wed Jul 13 15:50:42 2022

@author: Gilbr@
"""

from fonctions import *

# If this module is not called
if __name__ == '__main__':
    
    # If the different payment proposals are some things
    if (payementProposalTmbAutresCDF or \
        payementProposalTmbAutresUSD or \
        payementProposalTmbTmbCDF or \
        payementProposalTmbTmbUSD or \
        allPayementProposal):
        # Inform the user that we begins
        logger.info('Running...')
        # Call all functions
        main()
        # So inform that it's finished
        logger.info('End of proccess !')
        # Close the log file
        logHandler.close()    
        # We make a popup information
        toaster = win10toast.ToastNotifier()
        toaster.show_toast(
            "Conversion finished",
            f"Outputs are stored in {basePath}Output_TextFiles/\
                \nCheck {basePath}/Utils/log_messages.log for more details.",
            duration=10,
            icon_path=basePath + "prog/Bconvert.ico"
            )