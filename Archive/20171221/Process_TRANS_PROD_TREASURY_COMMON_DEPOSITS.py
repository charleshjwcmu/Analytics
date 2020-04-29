# -*- coding: utf-8 -*-
"""
Created on Wed Nov  1 16:01:25 2017
Process treasury common deposits file
@author: e620927
"""
PRODUCTION_ENVRIONMENT = True
################# Global Parameters  #################
if PRODUCTION_ENVRIONMENT:
    code_dir = "Z:/Charles/ORM/SourceCodes-Production"
    output_dir = "Z:/Charles/ORM/Output_Common"
    pdf_dir = "Z:/Charles/ORM/PDFReport_Common"
    mycommontable_dir = "Z:/FTDRDataBase/OPS_DEPOSIT_PBUP_FBO_DDA_GTRM"
    input_file = "Z:/Charles/ORM/Data/Parameters.xlsx"
    myuptable_dir = "Z:/FTDRDataBase/PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM"
    date_mapping_file = "Z:/Charles/ORM/Data/Dates_Mapping.xlsx"
    misc_dir = "Z:/FTDRDataBase/MISC_TABLES"
    days_need_update = 5
else:    
    code_dir = "Z:/Charles/ORM/SourceCodes"
    output_dir = "Z:/Charles/ORM/Output_Test"
    pdf_dir = "Z:/Charles/ORM/PDFReport_Test"
    mycommontable_dir = "Z:/FTDRDataBase/OPS_DEPOSIT_PBUP_FBO_DDA_GTRM"
    input_file = "Z:/Charles/ORM/Data/Parameters.xlsx"
    myuptable_dir = "Z:/FTDRDataBase/PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM"
    date_mapping_file = "Z:/Charles/ORM/Data/Dates_Mapping.xlsx"
    misc_dir = "Z:/FTDRDataBase/MISC_TABLES"
    days_need_update = 5

#daily_operational_balance_input_file = "Z:/Charles/ORM/Data/DAILY_OPERATIONAL_BALANCE_SUMMARY_20171129.xlsx"
################# Load Libraries #################
import os
import pandas as pd
import time
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import datetime
#import subprocess
#import shutil
import imp
import gc

os.chdir(code_dir)
Balance_Type = "Spot" # or "Availale"
import UpdateDatabase_DB
from UpdateDatabase_DB import read_db
imp.reload(UpdateDatabase_DB)
from Functions_Analysis import summarizeTableByAttribute, summarizeTable, formatDate

#################### Configuration  ####################
# read dates of input files
date_mapping = pd.ExcelFile(date_mapping_file).parse("Sheet1")
date_mapping.set_index("table_name",inplace=True,drop=True)
date_mapping = date_mapping.dropna()

# create folder hierarchy
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
if not os.path.exists(pdf_dir):
    os.makedirs(pdf_dir)
if not os.path.exists(myuptable_dir):
    os.makedirs(myuptable_dir)    
if not os.path.exists(myuptable_dir+"/RawData"):
    os.makedirs(myuptable_dir+"/RawData")
if not os.path.exists(mycommontable_dir):
    os.makedirs(mycommontable_dir)
if not os.path.exists(mycommontable_dir+"/RawData"):
    os.makedirs(mycommontable_dir+"/RawData")
if not os.path.exists(misc_dir):
    os.makedirs(misc_dir)    
if not os.path.exists(misc_dir+"/RawData"):
    os.makedirs(misc_dir+"/RawData")
    os.makedirs(misc_dir+"/HdfData")
    
day = datetime.date.today().strftime("%Y%m%d")
plot_folder = output_dir+"/Output_ORM_"+day
if not os.path.exists(plot_folder):
    os.mkdir(plot_folder)
#else:
#    shutil.rmtree(plot_folder)
#    os.makedirs(plot_folder)

pd.options.mode.chained_assignment = None
plt.style.use('ggplot')

styles = ['bs-','go-','r^-','cs-','mo-','y^-','kp-']
styles2 = ['bs--','go--','r^--','cs--','mo--','y^--','kp--']
styles3 = ['bh--','gh--','rh--','ch--','mh--','yh--','kh--']
def percentages(x):
    return '{:.3%}'.format(x)
def millions(x, pos=0):
    return '${:,.1f}MM'.format(x*1e-6)
def billions(x, pos=0):
    return '${:,.1f}B'.format(x*1e-9)
formatter_millions = FuncFormatter(millions)
formatter_billions = FuncFormatter(billions)
#################### Finish Configuration  ####################

def execute_ORM_FBO(dates, Verbose = False):
    print("############### Execute GTRM ORM Model: " + formatDate(dates.name))
    file_name_vw = mycommontable_dir+"/RawData/OPS_DEPOSIT_PBUP_FBO_DDA_GTRM_"+formatDate(dates.name)+".xlsx"
    file_name_vw_hdf = mycommontable_dir+"/HdfData/OPS_DEPOSIT_PBUP_FBO_DDA_GTRM_"+formatDate(dates.name)+".hdf"
    file_name_up = myuptable_dir+"/RawData/PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM_"+formatDate(dates.name)+".xlsx"
    if os.path.exists(file_name_vw) and os.path.exists(file_name_up) and Verbose:
        print("Already Exists")
        return()
#    required_tables = ["TRANS_PROD_TREASURY_COMMON_DEPOSITS","TBL_REF_PRODUCT_LOOKUP","TBL_REF_HEDGE_FUND","REF_OPSDEP_CATEGORY","PBUP_OPSDEP_DDA_XREF","REF_CDMS_CUSTOMER","PBUP_OPSDEP_RAI_OUTPUT"]
    #################### Read Data Table ##################
    start_time = time.time()
    
    table_name = "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP"
    data_up = read_db(table_name,formatDate(dates[table_name]))
    
#    table_name = "PBUP_OPSDEP_FUNDBAL_OUTPUT_DDA"
#    data_dda = read_db(table_name,formatDate(dates[table_name]))
    
    table_name = "TBL_REF_PRODUCT_LOOKUP"                #product exclusion list
    ref_product = read_db(table_name,formatDate(dates[table_name]))
    if len(ref_product)==0:
        print("Cannot find table: "+table_name)
        return()
    ref_product = ref_product[ref_product["EXCLUDES"]=="Y"]
        
    table_name = "TBL_REF_HEDGE_FUND"                   #hedge fund exclusion list
    ref_hedgefund = read_db(table_name,formatDate(dates[table_name]))
    if len(ref_hedgefund)==0:
        print("Cannot find table: "+table_name)
        return()

    table_name = "REF_OPSDEP_CATEGORY"                  #non discretionary fund list
    ref = read_db(table_name,formatDate(dates[table_name]))
    if len(ref)==0:
        print("Cannot find table: "+table_name)
        return()
    
    table_name = "PBUP_OPSDEP_DDA_XREF"                 #XRef table that has region and ult parent at DDA level
    mapping_xref = read_db(table_name,formatDate(dates[table_name]))
    if len(mapping_xref)==0:
        print("Cannot find table: "+table_name)
        return()
    mapping_xref.drop(["CREATED_ON"], axis=1, inplace=True)
    
    table_name = "REF_CDMS_CUSTOMER"                    #CDMS Customer table
    mapping_customer = read_db(table_name,formatDate(dates[table_name]))
    if len(mapping_customer)==0:
        print("Cannot find table: "+table_name)
        return()
    mapping_parentidname = mapping_customer[["ULT_PARENT_CD","ULT_PARENT_NAME"]].drop_duplicates()
#    mapping_customeridname = mapping_customer[["CUST_CD","CUST_NAME"]].drop_duplicates()
    
    table_name = "OPSDEP_MASTERATTR"                    #OPSDEP_MASTERATTR table - style
    mapping_attribute = read_db(table_name,formatDate(dates[table_name]))
    if len(mapping_attribute)==0:
        print("Cannot find table: "+table_name)
        return()
    mapping_attribute = mapping_attribute[["FUND_ID","FUND_NAME","DIVISION_GROUP","DIVISION","GROUP_MARKET_SEGMENT","MARKET_SEGMENT_DESC","STYLE","STYLE_GROUP"]]

    table_name = "PBUP_OPSDEP_RAI_OUTPUT"
    rai = read_db(table_name,formatDate(dates[table_name]))
    if len(rai)==0:
        print("Cannot find table: "+table_name)
        return()    
    rai.set_index(['ULT_PARENT_CD', 'REGION', 'ND_IND'], inplace=True, drop=False)
    #calculate binding constraint for RAI file
    bindings = rai["PERC_BAL_USD_75"].copy()
    bindings.name = "binding_constraint"
    select = rai["BEHAVIORAL_GROUP"] == "INTRA-DAY"
    bindings[select] = rai.loc[select,"UPPER_BOUND_1"]
    select = rai["BEHAVIORAL_GROUP"] != "INTRA-DAY"
    bindings[select] = rai.loc[select,["UPPER_BOUND_1","UPPER_BOUND_2"]].apply(min,axis=1)
    rai = pd.concat([rai,bindings],axis=1)
    
    table_name = "TRANS_PROD_TREASURY_COMMON_DEPOSITS"    #treasury common table
    #base = datetime.datetime.today()
    #date_list = [(base - datetime.timedelta(days=x)).strftime("%Y%m%d") for x in range(0, 100)]
    #data_common = read_db(table_name,["20171020","20171023","20171024","20171025"])
    data_common = read_db(table_name,formatDate(dates[table_name]))
    if len(data_common)==0:
        print("Cannot find table: "+table_name)
        return()

    #################### Clean and Filter Treasury Common Table ##################
    data_common_clean = data_common
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
    if Verbose:
        print("Writing accounts eliminated because of product exclusion list")
        product_exclusion.to_excel(plot_folder+"/exclusion_product"+formatDate(dates.name)+".xlsx")
    
    data_common_clean = data_common_clean[~(data_common_clean["PRIN_PARENT_NODE"].isin(ref_product["ACCT_HIER_SEQ_CODE"])&data_common_clean["IFS_PRODUCT"].isin(ref_product["IFS_PRODUCT"]))]
    print(sum(data_common_clean["PRINCIPAL_BAL_USD"]))
    
    # filter: exclude hedge fund list: daily availale balance
    if Balance_Type == "Available":
        data_common_clean = data_common_clean[~data_common_clean["FUND_NUM"].isin(ref_hedgefund["FUND_NUMBER"])]
        print(sum(data_common_clean["PRINCIPAL_BAL_USD"]))
    
#    if Verbose:
#        print("Writing daily clean common table")
#        data_common_clean.to_excel(plot_folder+"/data_common_clean_"+formatDate(dates.name)+".xlsx")
    
    #################### Enrich table with region and ultimate parent code ##################
    # pick up ND flag
    data_common_clean["ND_IND"] = ["N"]*len(data_common_clean.index)
    data_common_clean.loc[data_common_clean["SOURCE_CONTRACT_NUM"].isin(ref["SOURCE_CONTRACT_NUM"]),"ND_IND"] = "Y"
    
    # pick up region and ult parents
    priority1 = data_common_clean["FUND_NUM"].isin(mapping_xref["FUND_NUM"])
    priority2 = ~data_common_clean["FUND_NUM"].isin(mapping_xref["FUND_NUM"])&data_common_clean["SOURCE_CONTRACT_NUM"].isin(mapping_xref["DDA"])
    priority3 = ~data_common_clean["FUND_NUM"].isin(mapping_xref["FUND_NUM"])&~data_common_clean["SOURCE_CONTRACT_NUM"].isin(mapping_xref["DDA"])
    data_common_clean_p1 = data_common_clean[priority1]
    data_common_clean_p2 = data_common_clean[priority2]
    data_common_clean_p3 = data_common_clean[priority3]
    
    mapping_xref_tmp = mapping_xref[["FUND_NUM","REGION","ULT_PARENT_CD"]]
    mapping_xref_tmp = mapping_xref_tmp.drop_duplicates()
    data_common_clean_p1 = data_common_clean_p1.merge(mapping_xref_tmp,left_on='FUND_NUM',right_on='FUND_NUM', how='left')
    
    mapping_xref_tmp = mapping_xref[["DDA","REGION","ULT_PARENT_CD"]]
    mapping_xref_tmp = mapping_xref_tmp.drop_duplicates()
    data_common_clean_p2 = data_common_clean_p2.merge(mapping_xref_tmp,left_on='SOURCE_CONTRACT_NUM',right_on='DDA', how='left')
    del data_common_clean_p2["DDA"]
        
    data_common_clean = data_common_clean_p1.append(data_common_clean_p2)
    
    if Balance_Type == "Spot":
        data_common_clean_p3 = data_common_clean_p3.merge(mapping_xref_tmp,left_on='SOURCE_CONTRACT_NUM',right_on='DDA', how='left')
        del data_common_clean_p3["DDA"]
        data_common_clean_p3['ULT_PARENT_CD']="UnregulatedFund"
        data_common_clean = data_common_clean.append(data_common_clean_p3)
        data_common_clean = data_common_clean[~(data_common_clean["PRIN_PARENT_NODE"].isin(ref_product["ACCT_HIER_SEQ_CODE"])&data_common_clean["IFS_PRODUCT"].isin(ref_product["IFS_PRODUCT"]))]

    #################### Agg DDA to Parent/region/ND_flag level ##################
    table_total_deposit = summarizeTable(data_common_clean,["ULT_PARENT_CD","REGION","ND_IND"],"PRINCIPAL_BAL_USD")
    
    #################### calculate operational deposit and Output file ##################
    #################### Ultimate Parent Level
    dates_all = table_total_deposit.columns
    
    table_total_deposit_rai = table_total_deposit.join(rai,how='left',rsuffix="rai")
    for j in range(len(dates_all)):
    #    j = 0
        operational_deposit = table_total_deposit_rai[dates_all[j]].copy()
        
        select = operational_deposit > table_total_deposit_rai["binding_constraint"]
        operational_deposit[select] = table_total_deposit_rai.loc[select,"binding_constraint"]
        
    #    table_operational_deposit[dates_all[j]] = operational_deposit
        operational_deposit = operational_deposit.to_frame("OPERATIONAL_BAL_USD")
        operational_deposit["DAILY_SPOT_BAL_USD"] = table_total_deposit_rai[dates_all[j]]
        
        operational_deposit['AS_OF_DATE'] = dates_all[j]
        if j == 0:
            table_operational_deposit = operational_deposit
        else:
            table_operational_deposit = table_operational_deposit.append(operational_deposit)
    
    table_operational_deposit_summary = summarizeTableByAttribute(table_operational_deposit,["DAILY_SPOT_BAL_USD","OPERATIONAL_BAL_USD"]).copy()
    table_operational_deposit_summary.applymap(billions).to_excel(plot_folder+"/operational_deposit_by_dates"+formatDate(dates.name)+".xlsx")
    
    table_operational_deposit = table_operational_deposit.join(rai,how="left")
    table_operational_deposit = table_operational_deposit.merge(mapping_parentidname,left_on="ULT_PARENT_CD",right_on="ULT_PARENT_CD",how="left")
    table_operational_deposit.set_index(["ULT_PARENT_CD","REGION","ND_IND"],inplace=True,drop=False)
    #table_operational_deposit.head()
    
    table_operational_deposit['CREATED_ON'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    perc = table_operational_deposit["OPERATIONAL_BAL_USD"]/table_operational_deposit["DAILY_SPOT_BAL_USD"]
    perc[abs(table_operational_deposit["DAILY_SPOT_BAL_USD"])<0.00001] = 0
    table_operational_deposit["SPOT_BAL_PERC"] = perc
    
    table_operational_deposit = table_operational_deposit.sort_values(["AS_OF_DATE","ULT_PARENT_CD","REGION","ND_IND"])
    
    #table_operational_deposit[table_operational_deposit["AS_OF_DATE"] == "2017-10-24"].to_excel(plot_folder+"/holistictable.xlsx",merge_cells=False)
    if Verbose and not os.path.exists(file_name_up):
        print("Writing GTRM execution results on UP level")
        table_operational_deposit_excel = table_operational_deposit.copy()
        table_operational_deposit_excel.index = table_operational_deposit_excel["ULT_PARENT_CD"]+table_operational_deposit_excel["REGION"]+table_operational_deposit_excel["ND_IND"]
        table_operational_deposit_excel.to_excel(file_name_up,merge_cells=False)
    if len(data_up)!=0:
        data_up.set_index(["ULT_PARENT_CD","REGION","ND_IND"],inplace=True,drop=False)
        table_compare_up = table_operational_deposit.join(data_up,how="left",rsuffix="up")
        #table_compare_up.to_excel(plot_folder+"/compare_to_up.xlsx")
        if not all(abs(table_compare_up["OPERATIONAL_BAL_USD"] == table_compare_up["OPERATIONAL_BAL_USDup"]) <= 1):
            print("Warning: UP level calculation cannot be compared to FTDR UP level report.")
    
    #################### DDA level
    # get percentage at UP level
    table_operational_deposit["AS_OF_DATE"] = table_operational_deposit["AS_OF_DATE"].map(str)
    
    percentage = table_operational_deposit[["ULT_PARENT_CD","BEHAVIORAL_GROUP","ULT_PARENT_NAME","REGION","ND_IND","binding_constraint","SPOT_BAL_PERC"]].copy()
    percentage["id"] = table_operational_deposit[["ULT_PARENT_CD","REGION","ND_IND","AS_OF_DATE"]].apply(lambda x: ''.join(map(str, x)),axis=1)
    del percentage["ULT_PARENT_CD"]
    del percentage["REGION"]
    del percentage["ND_IND"]

    data_common_clean["AS_OF_DATE"] = data_common_clean["AS_OF_DATE"].map(str)
    ids = data_common_clean[["ULT_PARENT_CD","REGION","ND_IND","AS_OF_DATE"]].copy()
    ids = ids.replace(np.NaN, 'Missing')
    data_common_clean["id"] = ids.apply(lambda x: ''.join(x),axis=1)
    
    table_operational_deposit_DDA = data_common_clean.merge(percentage,left_on='id',right_on='id',suffixes=('','_y'), how='left')
    table_operational_deposit_DDA['OPERATIONAL_BAL_USD'] = table_operational_deposit_DDA["PRINCIPAL_BAL_USD"]*table_operational_deposit_DDA["SPOT_BAL_PERC"]
    table_operational_deposit_DDA['DAILY_EXCESS_BAL_USD'] = table_operational_deposit_DDA["PRINCIPAL_BAL_USD"] - table_operational_deposit_DDA['OPERATIONAL_BAL_USD']
#    summarizeTableByAttribute(table_operational_deposit_DDA,["PRINCIPAL_BAL_USD","OPERATIONAL_BAL_USD"]).applymap(billions)
    
    table_operational_deposit_DDA.rename(columns=
                                         {'CDMS_COUNTERPARTY_NAME': "CDMS_CUST_CODE",\
                                        'CDMS_CUST_NAME': 'CDMS_COUNTERPARTY_NAME',\
                                        'OPERATIONAL_BAL_USD':'DAILY_OPERATIONAL_BAL_USD',\
                                        'binding_constraint': 'OPERATIONAL_BINDING_CONSTRAINT',\
                                        'PRINCIPAL_BAL_USD': 'DAILY_SPOT_BAL_USD'},inplace=True)
    table_operational_deposit_DDA = table_operational_deposit_DDA.merge(mapping_attribute,left_on="FUND_NUM",right_on="FUND_ID",how="left")
#    import collections
#    [item for item, count in collections.Counter(data_common_clean.columns).items() if count > 1]

    if Verbose and not os.path.exists(file_name_vw):
        print("Writing GTRM execution results on DDA level")
        table_operational_deposit_DDA.to_excel(file_name_vw)
        table_operational_deposit_DDA.to_hdf(file_name_vw_hdf,"w",table=True)
    
    print("###############" + str(round(time.time()-start_time,0)) + " seconds\n")
    return(table_operational_deposit_summary)

################# Execution #################
##### execute ORM model with newly available common tables
orm_summary = pd.DataFrame()
#   days_need_update = 20
update_start = (datetime.datetime.today() - datetime.timedelta(days=days_need_update)).strftime('%Y-%m-%d')
update_end = datetime.datetime.today().strftime('%Y-%m-%d')

start= [i for i,x in enumerate(date_mapping.index==update_start) if x][0]
end = [i for i,x in enumerate(date_mapping.index==update_end) if x][0]

for i in range(start,end+1):
#for i in range(19,25):
#    i = 59
    dates = date_mapping.iloc[i,:]
    a = execute_ORM_FBO(dates,Verbose=True)
    if len(a) != 0:
        orm_summary = orm_summary.append(a)
orm_summary.to_excel(plot_folder+"/ORM_Summary.xlsx")

###### convert excels to hdf db
#imp.reload(UpdateDatabase_DB)

##### Daily Balance data
def update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type,days_need_update=150):

    if not os.path.exists(file_total_deposit):
        Balance_History = pd.DataFrame()
        Balance_History_client = pd.DataFrame()
        for i in range(len(date_mapping)):
            date = date_mapping.index[i].strftime("%Y%m%d")
            table_name = "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM"
            table_tmp = read_db(table_name,date)
            table_tmp = table_tmp[~pd.isnull(table_tmp.index)]

            if len(table_tmp) >0:
                balance = table_tmp[[balance_type]]
                date = table_tmp['AS_OF_DATE'].drop_duplicates().values
                balance.columns = date
                if len(Balance_History) == 0:
                    Balance_History = balance
                else:
                    Balance_History = Balance_History.join(balance,how="outer")
                    
                balance = summarizeTable(table_tmp,'ULT_PARENT_CD',balance_type,fillna=True)
                balance.columns = date
                
                if len(Balance_History_client) == 0:
                    Balance_History_client = balance
                else:
                    Balance_History_client = Balance_History_client.join(balance,how="outer")
                table_tmp = []
                gc.collect()
        Balance_History[~pd.isnull(Balance_History.index)].to_excel(file_total_deposit)
        Balance_History_client[~pd.isnull(Balance_History_client.index)].to_excel(file_total_deposit_client)
    else:
        Balance_History = pd.ExcelFile(file_total_deposit).parse("Sheet1")
        Balance_History_client = pd.ExcelFile(file_total_deposit_client).parse("Sheet1")
        Balance_History_client = Balance_History_client.set_index(['ULT_PARENT_CD'])
    
        update_start = (datetime.datetime.today() - datetime.timedelta(days=days_need_update)).strftime('%Y-%m-%d')
        update_end = datetime.datetime.today().strftime('%Y-%m-%d')

        start= [i for i,x in enumerate(date_mapping.index==update_start) if x][0]
        end = [i for i,x in enumerate(date_mapping.index==update_end) if x][0]
           
        for i in range(start,end+1):
            if not any(date_mapping.index[i]==Balance_History.columns):
                date = date_mapping.index[i].strftime("%Y%m%d")
                table_name = "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM"
                table_tmp = read_db(table_name,date)
                table_tmp = table_tmp[~pd.isnull(table_tmp.index)]
                if len(table_tmp) >0:
                    balance = table_tmp[[balance_type]]
                    date = table_tmp['AS_OF_DATE'].drop_duplicates().values
                    balance.columns = date
                    Balance_History = Balance_History.join(balance,how="outer")
                    table_tmp = []
                    gc.collect()
            if not any(date_mapping.index[i]==Balance_History_client.columns):
                date = date_mapping.index[i].strftime("%Y%m%d")
                table_name = "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM"
                table_tmp = read_db(table_name,date)
                table_tmp = table_tmp[~pd.isnull(table_tmp.index)]
                if len(table_tmp) >0:
                    balance = summarizeTable(table_tmp,'ULT_PARENT_CD',balance_type,fillna=True)
                    date = table_tmp['AS_OF_DATE'].drop_duplicates().values
                    balance.columns = date
                    Balance_History_client = Balance_History_client.join(balance,how="outer")
                    table_tmp = []
                    gc.collect()
    
        Balance_History[~pd.isnull(Balance_History.index)].to_excel(file_total_deposit)
        Balance_History_client[~pd.isnull(Balance_History_client.index)].to_excel(file_total_deposit_client)

balance_type = "DAILY_SPOT_BAL_USD"
file_total_deposit = misc_dir+"/RawData/TotalDepositTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/TotalDepositTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)
##### Daily Balance data
balance_type = "OPERATIONAL_BAL_USD"
file_total_deposit = misc_dir+"/RawData/OperationalDepositTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/OperationalDepositTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)

balance_type = "binding_constraint"
file_total_deposit = misc_dir+"/RawData/BindingConstraintTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/BindingConstraintTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)

balance_type = "BEHAVIORAL_GROUP"
file_total_deposit = misc_dir+"/RawData/BehaviorGroupTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/BehaviorGroupTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)

balance_type = "COHORT_RATIO"
file_total_deposit = misc_dir+"/RawData/CohortRatioTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/CohortRatioTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)

balance_type = "AVG_PAYMENT"
file_total_deposit = misc_dir+"/RawData/AveragePaymentTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/AveragePaymentTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)

