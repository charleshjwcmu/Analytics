# -*- coding: utf-8 -*-
"""
Created on Tue May 22 16:58:11 2018

@author: e620927
"""
PRODUCTION_ENVRIONMENT = True

print("Execute PREPROC Preprocess")

if PRODUCTION_ENVRIONMENT:
    code_dir = "Z:/Charles/ORM/SourceCodes-Production"
    output_dir = "Z:/Charles/ORM/Output_Common"
else:
    code_dir = "Z:/Charles/ORM/SourceCodes"
    output_dir = "Z:/Charles/ORM/Output_Test"
    
misc_dir = "Z:/FTDRDataBase/MISC_TABLES"
date_mapping_file = "Z:/Charles/ORM/Data/Dates_Mapping.xlsx"
table_name_up_trx_gtrm = "TBL_DEPOSIT_TRX_UP_GTRM"
gtrmtrxtable_dir = "Z:/FTDRDataBase/"+table_name_up_trx_gtrm #UP level, GTRM's RAI

################# Load Libraries #################
import os
import pandas as pd
import time
import datetime
import imp
import gc

os.chdir(code_dir)
import UpdateDatabase_DB
imp.reload(UpdateDatabase_DB)
from UpdateDatabase_DB import read_db

import Functions_Analysis
imp.reload(Functions_Analysis)
from Functions_Analysis import formatDate

#################### Configuration  ####################
# read dates of input files
date_mapping = pd.ExcelFile(date_mapping_file).parse("Sheet1")
date_mapping.set_index("table_name",inplace=True,drop=True)
date_mapping = date_mapping.dropna()
if not os.path.exists(gtrmtrxtable_dir):
    os.makedirs(gtrmtrxtable_dir)
    os.makedirs(gtrmtrxtable_dir+"/RawData")
    os.makedirs(gtrmtrxtable_dir+"/HdfData")

################### Clean up Historical Results ###############
start_time = time.time()

payment_table = "Payment_By_ParentRegion"
filename_preproc = misc_dir+"/RawData/"+payment_table+".xlsx"
table_name = "PBUP_OPSDEP_PREPROC"
preproc_files = date_mapping[table_name].drop_duplicates()

if not os.path.exists(filename_preproc):
    preproc_final = pd.DataFrame()
else:
    preproc_final = read_db(payment_table)
    preproc_files = preproc_files[preproc_files - datetime.timedelta(days=25) > max(preproc_final.columns)]

if len(preproc_files) > 0:
    for preproc_file in preproc_files:
    #        preproc_file = preproc_files[0]
        preproc = read_db(table_name,formatDate(preproc_file))
        
        if len(preproc) >0:
            preproc['id']=preproc["ULT_PARENT_CD"]+preproc["REGION"]+preproc["ND_IND"]
            preproc = preproc.drop(['AS_OF_PERIOD','CREATED_BY','CREATED_ON','PERIOD_ID','STYLE','TXN_CURRENCY','UPDATED_BY','UPDATED_ON',"ULT_PARENT_CD","REGION","ND_IND"],axis=1)
            preproc.rename(columns={"﻿AS_OF_DATE":u"AS_OF_DATE"},inplace=True)
            preproc = preproc.pivot(index='id',columns='AS_OF_DATE',values='TOTAL_OUTFLOW')
            preproc.columns=[datetime.datetime.strptime(x,'%m/%d/%Y %H:%M:%S AM')-datetime.timedelta(days=0.5) for x in preproc.columns]
            preproc = preproc.transpose().sort_index().transpose()
            
            if len(preproc_final) == 0:
                preproc_final = preproc.copy()
            else:
                previous = preproc_final.loc[:,~preproc_final.columns.isin(preproc.columns)]
                preproc_final = previous.join(preproc,how="right")
            preproc=[]
        else:
            print("Error: Cannot find the Preproc file.")
    preproc_final.to_excel(filename_preproc)

payment_table = "Balance_By_ParentRegion"
filename_preproc = misc_dir+"/RawData/"+payment_table+".xlsx"
table_name = "PBUP_OPSDEP_PREPROC"
preproc_files = date_mapping[table_name].drop_duplicates()

if not os.path.exists(filename_preproc):
    preproc_final = pd.DataFrame()
else:
    preproc_final = read_db(payment_table)
    preproc_files = preproc_files[preproc_files - datetime.timedelta(days=25) > max(preproc_final.columns)]

if len(preproc_files) > 0:
    for preproc_file in preproc_files:
    #        preproc_file = preproc_files[0]
        preproc = read_db(table_name,formatDate(preproc_file))
        
        if len(preproc) >0:
            preproc['id']=preproc["ULT_PARENT_CD"]+preproc["REGION"]+preproc["ND_IND"]
            preproc = preproc.drop(['AS_OF_PERIOD','CREATED_BY','CREATED_ON','PERIOD_ID','STYLE','TXN_CURRENCY','UPDATED_BY','UPDATED_ON',"ULT_PARENT_CD","REGION","ND_IND"],axis=1)
            preproc.rename(columns={"﻿AS_OF_DATE":u"AS_OF_DATE"},inplace=True)
            preproc = preproc.pivot(index='id',columns='AS_OF_DATE',values='TOTAL_PRIN_BAL_USD')
            preproc.columns=[datetime.datetime.strptime(x,'%m/%d/%Y %H:%M:%S AM')-datetime.timedelta(days=0.5) for x in preproc.columns]
            preproc = preproc.transpose().sort_index().transpose()
            
            if len(preproc_final) == 0:
                preproc_final = preproc.copy()
            else:
                previous = preproc_final.loc[:,~preproc_final.columns.isin(preproc.columns)]
                preproc_final = previous.join(preproc,how="right")
            preproc=[]
        else:
            print("Error: Cannot find the Preproc file.")
    preproc_final.to_excel(filename_preproc)

preproc_final = []
gc.collect()
end_time = time.time()
print("PREPROC TABLE PROCESS Time is "+str(round(end_time-start_time,0))+" seconds")
