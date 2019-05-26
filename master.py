# -*- coding: utf-8 -*-
"""
@author: mk99999
"""
import Data_stream as DS
import TD_seq_calc as TD


if __name__ == "__main__":
    sender = ""
    recipients="" #use ; as delimiter
    markets = ["CFI2Z9"]
    
    '''Excel spreadsheet that contains historical data'''
    file_path=r""
    
    '''Excel spreadsheet that retrieves market updates'''
    df_path =r""
    
    '''If there is an update calculate TD seq and send mail, otherwise finish function'''
    if DS.update(markets, file_path, df_path):
        TD.td_start(markets, recipients, sender, file_path)
    else:
        message = "No update to data"
        recipients=""
        TD.send_mail(message, recipients, sender)