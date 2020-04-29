# -*- coding: utf-8 -*-
"""
Standalone process to extract FX rates from daily treasury common tables

@author: e620927
"""
PRODUCTION_ENVRIONMENT = True

print("Execute Common Table Preprocess")
import datetime
date_mapping_file = "Z:/Charles/ORM/Data/Dates_Mapping.xlsx"
code_dir = "Z:/Charles/ORM/SourceCodes-Production"
output_dir = "Z:/Charles/ORM/Output_Common"
table_name_fx = "FX_RATES"
fx_dir = "Z:/FTDRDataBase/"+table_name_fx #DDA level, FTDR's RAI
misc_dir = "Z:/FTDRDataBase/MISC_TABLES"

################# Load Libraries #################
import os
import pandas as pd
import time
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
if not os.path.exists(fx_dir):
    os.makedirs(fx_dir)
    os.makedirs(fx_dir+"/RawData")
    os.makedirs(fx_dir+"/HdfData")

################### Clean up Historical Results ###############
start_time = time.time()

def execute_ORM_FX(dates, Verbose = False):
    print("############### Execute FX rate extraction: " + formatDate(dates.name))
    filename_fx = fx_dir+"/RawData"+"/"+table_name_fx+"_"+formatDate(dates.name)+".xlsx"
    filename_fx_hdf = fx_dir+"/HdfData"+"/"+table_name_fx+"_"+formatDate(dates.name)+".hdf"
    if os.path.exists(filename_fx) and os.path.exists(filename_fx_hdf) and Verbose:
        print("Already Exists")
        return()

    #################### Read Data Table ##################
    start_time = time.time()
    
    table_name = "TRANS_PROD_TREASURY_COMMON_DEPOSITS"    #treasury common table
    data_common = read_db(table_name,formatDate(dates[table_name]))
    if len(data_common)==0:
        print("Cannot find table: "+table_name)
        return(None)

    #################### Clean and Filter Treasury Common Table ##################
    data_common_clean = data_common.copy()
#    data_common_clean = data_common

    fx_rates = data_common_clean[["TRANSACTION_CURRENCY","FX_BASE_TO_USD"]].drop_duplicates().set_index("TRANSACTION_CURRENCY").sort_index()
    fx_rates.columns=data_common_clean["AS_OF_DATE"].drop_duplicates().values
    fx_rates.to_excel(filename_fx)
    fx_rates.to_hdf(filename_fx_hdf,"w",table=True)

    print("###############" + str(round(time.time()-start_time,0)) + " seconds\n")
    return(fx_rates)

################# Execution #################
##### execute ORM model with newly available common tables
if PRODUCTION_ENVRIONMENT:
    days_need_update = 10
    update_start = (datetime.datetime.today() - datetime.timedelta(days=days_need_update)).strftime('%Y-%m-%d')
else:
    update_start = datetime.datetime.strptime("2017-04-26",'%Y-%m-%d')

    
update_end = datetime.datetime.today().strftime('%Y-%m-%d')
start= [i for i,x in enumerate(date_mapping.index==update_start) if x][0]
end = [i for i,x in enumerate(date_mapping.index==update_end) if x][0]

for i in range(start,end+1):
#for i in range(2,6):
#    i = 2
    dates = date_mapping.iloc[i,:]
    execute_ORM_FX(dates,Verbose=True)

file_name = misc_dir+"/RawData/"+table_name_fx+"_Summary.xlsx"
if os.path.exists(file_name):
    FX_summary = pd.ExcelFile(file_name).parse("Sheet1")
    FX_summary = FX_summary.set_index("TRANSACTION_CURRENCY")
    flag_change = False
    for i in range(start,end+1):
    #    i = 367
        if not any(date_mapping.index[i]==FX_summary.columns):
            data_fx = read_db(table_name_fx,formatDate(date_mapping.index[i]))
            if len(data_fx)!=0:
                if len(data_fx.index.unique()) != len(data_fx.index):
                    data_fx = data_fx[data_fx.iloc[:,0]!=1].append(data_fx.loc["USD",:]).sort_index()
                FX_summary = FX_summary.join(data_fx)
                flag_change = True
            data_fx = []
            gc.collect()
    if flag_change:
        FX_summary.transpose().sort_index().transpose().to_excel(file_name)
else:
    FX_summary = pd.DataFrame()
    for i in range(len(date_mapping.index)):
    #for i in range(2,5):
    #    i = 4
        data_fx = read_db(table_name_fx,formatDate(date_mapping.index[i]))
        if len(data_fx)!=0:
            if len(data_fx.index.unique()) != len(data_fx.index):
                data_fx = data_fx[data_fx.iloc[:,0]!=1].append(data_fx.loc["USD",:]).sort_index()
            if len(FX_summary) == 0:
                FX_summary = data_fx
            else:
                FX_summary = FX_summary.join(data_fx)
    FX_summary.to_excel(file_name)



