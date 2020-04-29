# -*- coding: utf-8 -*-
"""
Created on Fri May  4 14:44:32 2018
Balance by Currency
@author: e620927
"""
PRODUCTION_ENVRIONMENT = True

print("Execute Common Table Preprocess")
import datetime
date_mapping_file = "Z:/Charles/ORM/Data/Dates_Mapping.xlsx"
code_dir = "Z:/Charles/ORM/SourceCodes"
output_dir = "Z:/Charles/ORM/Output_Common"
table_name_core = "BALANCEBYCURRENCY"
table_name_core_usd = "BALANCEBYCURRENCY_USD"
core_dir = "Z:/FTDRDataBase/"+table_name_core #DDA level, FTDR's RAI
core_usd_dir = "Z:/FTDRDataBase/"+table_name_core_usd #DDA level, FTDR's RAI
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
from UpdateDatabase_DB import read_db,read_db_test

import Functions_Analysis
imp.reload(Functions_Analysis)
from Functions_Analysis import formatDate

#################### Configuration  ####################
# read dates of input files
date_mapping = pd.ExcelFile(date_mapping_file).parse("Sheet1")
date_mapping.set_index("table_name",inplace=True,drop=True)
date_mapping = date_mapping.dropna()
if not os.path.exists(core_dir):
    os.makedirs(core_dir)
    os.makedirs(core_dir+"/RawData")
    os.makedirs(core_dir+"/HdfData")
if not os.path.exists(core_usd_dir):
    os.makedirs(core_usd_dir)
    os.makedirs(core_usd_dir+"/RawData")
    os.makedirs(core_usd_dir+"/HdfData")
################### Clean up Historical Results ###############

def execute_Core_Balance(dates, Verbose = False):
    print("############### Execute Core Balance extraction: " + formatDate(dates.name))
    filename_core = core_dir+"/RawData"+"/"+table_name_core+"_"+formatDate(dates.name)+".xlsx"
    filename_core_hdf = core_dir+"/HdfData"+"/"+table_name_core+"_"+formatDate(dates.name)+".hdf"
    filename_core_usd = core_usd_dir+"/RawData"+"/"+table_name_core_usd+"_"+formatDate(dates.name)+".xlsx"
    filename_core_hdf_usd = core_usd_dir+"/HdfData"+"/"+table_name_core_usd+"_"+formatDate(dates.name)+".hdf"

    if os.path.exists(filename_core) and os.path.exists(filename_core_hdf) and Verbose:
        print("Already Exists")
        return()

    #################### Read Data Table ##################
    start_time = time.time()
    
    table_name = "TRANS_PROD_TREASURY_COMMON_DEPOSITS"    #treasury common table
    data_common = read_db_test(table_name,formatDate(dates[table_name]))
    if len(data_common)==0:
        print("Cannot find table: "+table_name)
        return()
    table_name = "TBL_REF_PRODUCT_LOOKUP"                #product exclusion list
    ref_product = read_db(table_name,formatDate(dates[table_name]))
    if len(ref_product)==0:
        print("Cannot find table: "+table_name)
        return()
    ref_product = ref_product[ref_product["EXCLUDES"]=="Y"]

    #################### Clean and Filter Treasury Common Table ##################
    data_common_clean = data_common.copy()
#    data_common_clean = data_common
    print(sum(data_common_clean["PRINCIPAL_BAL_USD"]))
    
    # filter: parent node begin with %B0
    nodes = data_common_clean["PRIN_PARENT_NODE"].tolist()
    filterone = [False]*len(nodes)
    for i in range(len(nodes)):
        if nodes[i][:3]=="%B0":
            filterone[i]=True
    data_common_clean = data_common_clean[filterone]
    print(sum(data_common_clean["PRINCIPAL_BAL_USD"]))
    
    # filter: TDR status = 0
    data_common_clean = data_common_clean[data_common_clean["TDR_STATUS_FLAG"]==0]
    print(sum(data_common_clean["PRINCIPAL_BAL_USD"]))
    
    # filter: internal transaction flag = 0 or 4
    data_common_clean = data_common_clean[data_common_clean["INTERNAL_COMPANY_FLAG"].isin([0,4])]
    print(sum(data_common_clean["PRINCIPAL_BAL_USD"]))
    
    # filter: exclude Non-Operational Product Exclusion List: daily spot balance
    product_exclusion = data_common_clean[(data_common_clean["PRIN_PARENT_NODE"].isin(ref_product["ACCT_HIER_SEQ_CODE"])&data_common_clean["IFS_PRODUCT"].isin(ref_product["IFS_PRODUCT"]))]
    print(sum(product_exclusion["PRINCIPAL_BAL_USD"]))
    data_common_clean = data_common_clean[~(data_common_clean["PRIN_PARENT_NODE"].isin(ref_product["ACCT_HIER_SEQ_CODE"])&data_common_clean["IFS_PRODUCT"].isin(ref_product["IFS_PRODUCT"]))]
    print("Available Balance is " + str(sum(data_common_clean["PRINCIPAL_BAL_USD"])))

    # filter: get MMIA and Foreign Time deposits
    flag = ref_product.loc[ref_product["PROD_GROUP"].isin(["MMIA","FOREIGN TIME"]),["ACCT_HIER_SEQ_CODE","IFS_PRODUCT"]]
    flag_product_exclusion = product_exclusion[(product_exclusion["PRIN_PARENT_NODE"].isin(flag["ACCT_HIER_SEQ_CODE"])&product_exclusion["IFS_PRODUCT"].isin(flag["IFS_PRODUCT"]))]
    print(sum(flag_product_exclusion["PRINCIPAL_BAL_USD"]))
    data_common_clean = data_common_clean.append(flag_product_exclusion)
    print(sum(data_common_clean["PRINCIPAL_BAL_USD"]))
    
    result_tmp = data_common_clean.groupby('TRANSACTION_CURRENCY')['PRINCIPAL_BAL'].sum()
    result_tmp = result_tmp.to_frame()
    result_tmp.columns = data_common_clean["AS_OF_DATE"].drop_duplicates()
    result_tmp.to_excel(filename_core)
    result_tmp.to_hdf(filename_core_hdf,"w",table=True)
    
    result_tmp_usd = data_common_clean.groupby('TRANSACTION_CURRENCY')['PRINCIPAL_BAL_USD'].sum()
    result_tmp_usd = result_tmp_usd.to_frame()
    result_tmp_usd.columns = data_common_clean["AS_OF_DATE"].drop_duplicates()
    result_tmp_usd.to_excel(filename_core_usd)
    result_tmp_usd.to_hdf(filename_core_hdf_usd,"w",table=True)
    print("###############" + str(round(time.time()-start_time,0)) + " seconds\n")
    return(result_tmp)

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
#for i in range(65,160):
#    i = 4
    dates = date_mapping.iloc[i,:]
    execute_Core_Balance(dates,Verbose=True)

    
for table_name in [table_name_core,table_name_core_usd]:
    file_name = misc_dir+"/RawData/"+table_name+"_Summary.xlsx"
    if os.path.exists(file_name):
        FX_summary = pd.ExcelFile(file_name).parse("Sheet1")
        FX_summary = FX_summary.set_index("TRANSACTION_CURRENCY")
        flag_change = False
        for i in range(start,end+1):
        #    i = 367
            if not any(date_mapping.index[i]==FX_summary.columns):
                data_fx = read_db(table_name,formatDate(date_mapping.index[i]))
                if len(data_fx)!=0:
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
            data_fx = read_db(table_name,formatDate(date_mapping.index[i]))
            if len(data_fx)!=0:
                if len(FX_summary) == 0:
                    FX_summary = data_fx
                else:
                    FX_summary = FX_summary.join(data_fx)
        FX_summary.to_excel(file_name)

