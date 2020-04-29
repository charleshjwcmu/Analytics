# -*- coding: utf-8 -*-
"""
Created on Wed Nov  1 16:01:25 2017
Process treasury common deposits file
@author: e620927
"""
print("Execute Common Table Preprocess")
import datetime
PRODUCTION_ENVRIONMENT = True

input_file = "Z:/Charles/ORM/Data/Parameters.xlsx"
date_mapping_file = "Z:/Charles/ORM/Data/Dates_Mapping.xlsx"
misc_dir = "Z:/FTDRDataBase/MISC_TABLES"
misc_dir_rai_gtrm = "Z:/FTDRDataBase/MISC_TABLES_RAI_GTRM"
table_name_dda_gtrm = "OPS_DEPOSIT_PBUP_FBO_DDA_GTRM"
table_name_up_gtrm = "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM"
table_name_dda_rai_gtrm = "OPS_DEPOSIT_DDA_RAI_GTRM"
table_name_up_rai_gtrm = "OPS_DEPOSIT_UP_RAI_GTRM"
table_name_up_trx_gtrm = "TBL_DEPOSIT_TRX_UP_GTRM"

mycommontable_dir = "Z:/FTDRDataBase/"+table_name_dda_gtrm #DDA level, FTDR's RAI
myuptable_dir = "Z:/FTDRDataBase/"+table_name_up_gtrm #UP level, FTDR's RAI
gtrmcommontable_dir = "Z:/FTDRDataBase/"+table_name_dda_rai_gtrm #DDA level, GTRM's RAI
gtrmuptable_dir = "Z:/FTDRDataBase/"+table_name_up_rai_gtrm #UP level, GTRM's RAI
gtrmtrxtable_dir = "Z:/FTDRDataBase/"+table_name_up_trx_gtrm #UP level, GTRM's RAI

################# Global Parameters  #################
if PRODUCTION_ENVRIONMENT:
    code_dir = "Z:/Charles/ORM/SourceCodes-Production"
    output_dir = "Z:/Charles/ORM/Output_Common"
    pdf_dir = "Z:/Charles/ORM/PDFReport_Common"
    COPY_TO_FOLDER = False
    dir_excel_copyto = "Z:/GTRM_IRRLiq_QA/LiqRisk/OpDepos/Reports/ParentLevel/ExcelOutput_Tests"
    days_need_update = 10
    update_start = (datetime.datetime.today() - datetime.timedelta(days=days_need_update)).strftime('%Y-%m-%d')
    UPDATE_RAI_GTRM = False
else:
    code_dir = "Z:/Charles/ORM/SourceCodes"
    output_dir = "Z:/Charles/ORM/Output_Test"
    pdf_dir = "Z:/Charles/ORM/PDFReport_Test"
    COPY_TO_FOLDER = False
    dir_excel_copyto = misc_dir
#    days_need_update = 240
#    update_start = datetime.datetime.strptime("2017-04-26",'%Y-%m-%d')
#    update_start = datetime.datetime.strptime("2017-08-10",'%Y-%m-%d')
    days_need_update = 10
    update_start = (datetime.datetime.today() - datetime.timedelta(days=days_need_update)).strftime('%Y-%m-%d')
    UPDATE_RAI_GTRM = False

#daily_operational_balance_input_file = "Z:/Charles/ORM/Data/DAILY_OPERATIONAL_BALANCE_SUMMARY_20171129.xlsx"
################# Load Libraries #################
import os
import pandas as pd
import time
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
#import subprocess
#import shutil
import imp
import gc

os.chdir(code_dir)
Balance_Type = "Spot" # or "Availale"
import UpdateDatabase_DB
imp.reload(UpdateDatabase_DB)
from UpdateDatabase_DB import read_db, delete_db

import Functions_Analysis
imp.reload(Functions_Analysis)
from Functions_Analysis import summarizeTableByAttribute, summarizeTable, formatDate, process_dda_GTRM

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
if not os.path.exists(myuptable_dir+"/HdfData"):
    os.makedirs(myuptable_dir+"/HdfData")
if not os.path.exists(mycommontable_dir):
    os.makedirs(mycommontable_dir)
if not os.path.exists(mycommontable_dir+"/RawData"):
    os.makedirs(mycommontable_dir+"/RawData")
if not os.path.exists(mycommontable_dir+"/HdfData"):
    os.makedirs(mycommontable_dir+"/HdfData")
if not os.path.exists(gtrmcommontable_dir):
    os.makedirs(gtrmcommontable_dir)
    os.makedirs(gtrmcommontable_dir+"/RawData")
    os.makedirs(gtrmcommontable_dir+"/HdfData")
if not os.path.exists(gtrmuptable_dir):
    os.makedirs(gtrmuptable_dir)
    os.makedirs(gtrmuptable_dir+"/RawData")
    os.makedirs(gtrmuptable_dir+"/HdfData")    
if not os.path.exists(gtrmtrxtable_dir):
    os.makedirs(gtrmtrxtable_dir)
    os.makedirs(gtrmtrxtable_dir+"/RawData")
    os.makedirs(gtrmtrxtable_dir+"/HdfData")
if not os.path.exists(misc_dir):
    os.makedirs(misc_dir)    
    os.makedirs(misc_dir+"/RawData")
if not os.path.exists(misc_dir_rai_gtrm):
    os.makedirs(misc_dir_rai_gtrm)
    os.makedirs(misc_dir_rai_gtrm+"/RawData")
    
day = datetime.date.today().strftime("%Y%m%d")
plot_folder = output_dir+"/Output_ORM_"+day
if not os.path.exists(plot_folder):
    os.mkdir(plot_folder)
#else:
#    shutil.rmtree(plot_folder)
#    os.makedirs(plot_folder)

plot_folder_gtrm = output_dir+"/Output_ORM_GTRM_"+day
if not os.path.exists(plot_folder_gtrm):
    os.mkdir(plot_folder_gtrm)

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

column_name_operational = "Operational_Deposits"
column_name_spot = "Total_Deposits"
column_name_excess = "Excess_Deposits"

column_name_company = "LegalEntity"
column_name_bu = "BusinessLines"
column_name_region = "Region"
column_name_group = "BehaviorGroup"
column_name_currency = "Currency"
column_name_parent = "ParentCompy"
column_name_parentname = 'Client'
column_name_style = "Style"
column_name_stylegroup = "FundStyle"
column_name_division = "BusinessUnit"
column_name_market = "MarketSegment"

column_name_productagg_type = "ProductSegmt"
column_name_product_type = "ProductType"
column_name_product_subtype = "ProdSubtype"
column_name_GL_type = "GLProdType"
#################### Finish Configuration  ####################

NEED_CLEANUP = False #if True, system will clean up historical analysis data between cleanup_start date and cleanup_end date
cleanup_start = datetime.datetime.strptime("2018-04-11",'%Y-%m-%d')
cleanup_end = datetime.datetime.today()

if NEED_CLEANUP:
    print("Start Cleanup")
    ##clean up DDA and UP file
    date_list = [(cleanup_end - datetime.timedelta(days=x)).strftime("%Y%m%d") for x in range((cleanup_end-cleanup_start).days+1)]
    
    table_name = "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM"
    delete_db(table_name,date_list)
    table_name = "OPS_DEPOSIT_PBUP_FBO_DDA_GTRM"
    delete_db(table_name,date_list)
    
    ##update summary files
    attributes = [column_name_division,column_name_stylegroup,'COUNTRY_DOMICILE','ProductTypes',column_name_currency,column_name_company,column_name_region,column_name_group,"ID",column_name_parent,"FUND_NUM"]
    balance_types = [column_name_spot,column_name_operational]
    file_names = []
    for attribute in attributes:
        for b_type in balance_types:
            file_names = file_names + [misc_dir+"/RawData/"+b_type+"_By_"+ attribute+".xlsx"]
    
    file_total_deposit = misc_dir+"/RawData/TotalDepositTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir+"/RawData/TotalDepositTableByParent.xlsx"
    file_names = file_names + [file_total_deposit,file_total_deposit_client]
    file_total_deposit = misc_dir+"/RawData/OperationalDepositTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir+"/RawData/OperationalDepositTableByParent.xlsx"
    file_names = file_names + [file_total_deposit,file_total_deposit_client]
    file_total_deposit = misc_dir+"/RawData/BindingConstraintTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir+"/RawData/BindingConstraintTableByParent.xlsx"
    file_names = file_names + [file_total_deposit,file_total_deposit_client]
    file_total_deposit = misc_dir+"/RawData/BehaviorGroupTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir+"/RawData/BehaviorGroupTableByParent.xlsx"
    file_names = file_names + [file_total_deposit,file_total_deposit_client]
    file_total_deposit = misc_dir+"/RawData/CohortRatioTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir+"/RawData/CohortRatioTableByParent.xlsx"
    file_names = file_names + [file_total_deposit,file_total_deposit_client]
    file_total_deposit = misc_dir+"/RawData/AveragePaymentTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir+"/RawData/AveragePaymentTableByParent.xlsx"
    file_names = file_names + [file_total_deposit,file_total_deposit_client]
    
    for file_name in file_names:
        if not os.path.exists(file_name):
            next
        print(file_name+": File Updated")
        data = pd.ExcelFile(file_name).parse('Sheet1',index_col=0)
        
        select = (data.columns >= cleanup_start)&(data.columns <= cleanup_end)
        data.ix[:,~select].to_excel(file_name)
    print("End cleanup")

################### Clean up Historical Results ###############
start_time = time.time()

def execute_ORM_FBO(dates, Verbose = False):
    print("############### Execute GTRM ORM Model: " + formatDate(dates.name))
    file_name_vw = mycommontable_dir+"/RawData/"+table_name_dda+"_"+formatDate(dates.name)+".xlsx"
    file_name_vw_hdf = mycommontable_dir+"/HdfData/"+table_name_dda+"_"+formatDate(dates.name)+".hdf"
    file_name_up = myuptable_dir+"/RawData/"+table_name_up+"_"+formatDate(dates.name)+".xlsx"
    file_name_up_hdf = myuptable_dir+"/HdfData/"+table_name_up+"_"+formatDate(dates.name)+".hdf"

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

    table_name = rai_table_name
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
    if Verbose:
        print("Writing accounts eliminated because of product exclusion list")
        product_exclusion.to_excel(plot_folder+"/exclusion_product"+formatDate(dates.name)+".xlsx")
    print(sum(product_exclusion["PRINCIPAL_BAL_USD"]))
    data_common_clean = data_common_clean[~(data_common_clean["PRIN_PARENT_NODE"].isin(ref_product["ACCT_HIER_SEQ_CODE"])&data_common_clean["IFS_PRODUCT"].isin(ref_product["IFS_PRODUCT"]))]
    print("Available Balance is " + str(millions(sum(data_common_clean["PRINCIPAL_BAL_USD"]))))

    # filter: get MMIA and Foreign Time deposits
    flag = ref_product.loc[ref_product["PROD_GROUP"].isin(["MMIA","FOREIGN TIME"]),["ACCT_HIER_SEQ_CODE","IFS_PRODUCT"]]
    flag_product_exclusion = product_exclusion[(product_exclusion["PRIN_PARENT_NODE"].isin(flag["ACCT_HIER_SEQ_CODE"])&product_exclusion["IFS_PRODUCT"].isin(flag["IFS_PRODUCT"]))]
    print(sum(flag_product_exclusion["PRINCIPAL_BAL_USD"]))
    
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
        # append MMIA and Foreign Time deposit back
        flag_product_exclusion["ND_IND"] = [np.nan]*len(flag_product_exclusion.index)
        flag_product_exclusion["REGION"] = [np.nan]*len(flag_product_exclusion.index)
        flag_product_exclusion["ULT_PARENT_CD"] = [np.nan]*len(flag_product_exclusion.index)
        data_common_clean = data_common_clean.append(flag_product_exclusion)
        print("Spot Balance is " + str(millions(sum(data_common_clean["PRINCIPAL_BAL_USD"]))))

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
    
    table_operational_deposit = table_operational_deposit.join(rai,how="left")
#    table_operational_deposit.head()
#    table_operational_deposit = 
#    table_operational_deposit.join(preproc,how="left")
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
        table_operational_deposit_excel.to_hdf(file_name_up_hdf,"w",table=True)
    if Verbose and len(data_up)!=0:
        data_up.set_index(["ULT_PARENT_CD","REGION","ND_IND"],inplace=True,drop=False)
        table_compare_up = table_operational_deposit[~(table_operational_deposit['ULT_PARENT_CD'].isnull()&table_operational_deposit['REGION'].isnull()&table_operational_deposit['ND_IND'].isnull())].join(data_up,how="left",rsuffix="up")
        if not all(abs(table_compare_up["OPERATIONAL_BAL_USD"] - table_compare_up["OPERATIONAL_BAL_USDup"]) <= 1):
            print("Warning: UP level calculation cannot be compared to FTDR UP level report.")
            table_compare_up.to_excel(plot_folder+"/DataNeedExamine_"+formatDate(dates.name)+".xlsx",merge_cells=False)

#        sum(table_operational_deposit[~(table_operational_deposit['ULT_PARENT_CD'].isnull()&table_operational_deposit['REGION'].isnull()&table_operational_deposit['ND_IND'].isnull())]["OPERATIONAL_BAL_USD"])
#        sum(table_operational_deposit["OPERATIONAL_BAL_USD"])
#        table_compare_up = data_up.join(table_operational_deposit[~(table_operational_deposit['ULT_PARENT_CD'].isnull()&table_operational_deposit['REGION'].isnull()&table_operational_deposit['ND_IND'].isnull())],how="left",rsuffix="up")
#        sum(table_compare_up["OPERATIONAL_BAL_USD"])

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
    
    table_operational_deposit_summary = summarizeTableByAttribute(table_operational_deposit_DDA,["DAILY_SPOT_BAL_USD","DAILY_OPERATIONAL_BAL_USD"]).copy()
    table_operational_deposit_summary.applymap(millions).to_excel(plot_folder+"/operational_deposit_by_dates"+formatDate(dates.name)+".xlsx")

    print("###############" + str(round(time.time()-start_time,0)) + " seconds\n")
    return(table_operational_deposit_summary)

################# Execution #################
##### execute ORM model with newly available common tables
orm_summary = pd.DataFrame()
rai_table_name = "PBUP_OPSDEP_RAI_OUTPUT"
table_name_dda = table_name_dda_gtrm
table_name_up = table_name_up_gtrm

update_end = datetime.datetime.today().strftime('%Y-%m-%d')

start= [i for i,x in enumerate(date_mapping.index==update_start) if x][0]
end = [i for i,x in enumerate(date_mapping.index==update_end) if x][0]

for i in range(start,end+1):
#for i in range(253,257):
#    i = 360
    dates = date_mapping.iloc[i,:]
    a = execute_ORM_FBO(dates,Verbose=True)
#    a = execute_ORM_FBO(dates,Verbose=False)

    if len(a) != 0:
        orm_summary = orm_summary.append(a)
#orm_summary.to_excel(plot_folder+"/ORM_Summary.xlsx")

################# Execution #################
##### execute ORM model with self calibrated RAI files
if UPDATE_RAI_GTRM:
    orm_summary = pd.DataFrame()
    
    rai_table_name = "PBUP_OPSDEP_RAI_GTRM"
    table_name_dda = table_name_dda_rai_gtrm
    table_name_up = table_name_up_rai_gtrm
    mycommontable_dir = gtrmcommontable_dir
    myuptable_dir = gtrmuptable_dir
    plot_folder = plot_folder_gtrm
    update_end = datetime.datetime.today().strftime('%Y-%m-%d')
    
    start= [i for i,x in enumerate(date_mapping.index==update_start) if x][0]
    end = [i for i,x in enumerate(date_mapping.index==update_end) if x][0]
    
    for i in range(start,end+1):
    #for i in range(253,257):
    #    i = 352
        dates = date_mapping.iloc[i,:]
        a = execute_ORM_FBO(dates,Verbose=True)
    #    a = execute_ORM_FBO(dates,Verbose=False)
    
        if len(a) != 0:
            orm_summary = orm_summary.append(a)
    orm_summary.to_excel(plot_folder+"/ORM_Summary_self_rai.xlsx")

####Transaction Data
#def execute_ORM_TRX(dates, Verbose = False):
#    print("############### Execute GTRM ORM Model: " + formatDate(dates.name))
#    file_name_trx = gtrmtrxtable_dir+"/RawData/"+table_name_up_trx_gtrm+"_"+formatDate(dates.name)+".xlsx"
#    file_name_trx_hdf = gtrmtrxtable_dir+"/HdfData/"+table_name_up_trx_gtrm+"_"+formatDate(dates.name)+".hdf"
#    if os.path.exists(file_name_trx) and os.path.exists(file_name_trx_hdf) and Verbose:
#        print("Already Exists")
#        return()
##    required_tables = ["TRANS_PROD_TREASURY_COMMON_DEPOSITS","TBL_REF_PRODUCT_LOOKUP","TBL_REF_HEDGE_FUND","REF_OPSDEP_CATEGORY","PBUP_OPSDEP_DDA_XREF","REF_CDMS_CUSTOMER","PBUP_OPSDEP_RAI_OUTPUT"]
#    #################### Read Data Table ##################
#    start_time = time.time()
#    table_name = "TBL_DEPOSIT_TRX"    #treasury common table
#    data_trx = read_db(table_name,formatDate(dates[table_name]))
#    if len(data_trx)==0:
#        print("Cannot find table: "+table_name)
#    else:
#        table_name = "REF_OPSDEP_CATEGORY"                  #non discretionary fund list
#        ref = read_db(table_name,formatDate(dates[table_name]))
#        if len(ref)==0:
#            print("Cannot find table: "+table_name)
#            return()
#    
#        table_name = "PBUP_OPSDEP_DDA_XREF"                 #XRef table that has region and ult parent at DDA level
#        mapping_xref = read_db(table_name,formatDate(dates[table_name]))
#        if len(mapping_xref)==0:
#            print("Cannot find table: "+table_name)
#            return()
#        mapping_xref.drop(["CREATED_ON"], axis=1, inplace=True)
#
#        types = data_trx['TRANSACTION_TYPE'].unique()
#        data_trx_result = pd.DataFrame()
#        for tran_type in types:
#            data_trx_filtered = data_trx[data_trx['TRANSACTION_TYPE']==tran_type]
#            data_trx_filtered = data_trx_filtered[['﻿ACCOUNT_NUMBER','AMOUNT_USD','AMOUNT','FUND_NUMBER','TRANSACTION_TYPE']]
#            mapping_xref_tmp = mapping_xref[["FUND_NUM","REGION","ULT_PARENT_CD"]]
#            mapping_xref_tmp = mapping_xref_tmp.drop_duplicates()
#    
#            data_trx_filtered = data_trx_filtered.merge(mapping_xref_tmp,left_on='FUND_NUMBER',right_on='FUND_NUM', how='left')
#            data_trx_filtered["ND_IND"] = ["N"]*len(data_trx_filtered.index)
#            data_trx_filtered.loc[data_trx_filtered["﻿ACCOUNT_NUMBER"].isin(ref["SOURCE_CONTRACT_NUM"]),"ND_IND"] = "Y"
#            data_trx_filtered = data_trx_filtered.groupby(['ULT_PARENT_CD','REGION','ND_IND']).sum()
#            data_trx_filtered['TRANSACTION_TYPE']=tran_type
#            data_trx_result = data_trx_result.append(data_trx_filtered)
#        data_trx_result['AS_OF_DATE'] = data_trx['AS_OF_DATE'].drop_duplicates()[0]
#        data_trx_result.to_excel(file_name_trx,merge_cells=False)
#        data_trx_result.to_hdf(file_name_trx_hdf,"w",table=True)
#    print("###############" + str(round(time.time()-start_time,0)) + " seconds\n")
#
#update_start_trx = (datetime.datetime.today() - datetime.timedelta(days=30)).strftime('%Y-%m-%d')
#update_end = datetime.datetime.today().strftime('%Y-%m-%d')
#start= [i for i,x in enumerate(date_mapping.index==update_start_trx) if x][0]
#end = [i for i,x in enumerate(date_mapping.index==update_end) if x][0]
#
#for i in range(start,end+1):
#    dates = date_mapping.iloc[i,:]
#    execute_ORM_TRX(dates,Verbose=True)
#
## Run history with PREPROC file as of certain period
#filename_preproc = misc_dir+"/RawData/"+"Payment"+"_By_"+ "ParentRegion"+".xlsx"
#if not os.path.exists(filename_preproc):
#    table_name = "PBUP_OPSDEP_PREPROC"
#    preproc = read_db(table_name,"20180411")
#    preproc['id']=preproc["ULT_PARENT_CD"]+preproc["REGION"]+preproc["ND_IND"]
#    preproc = preproc.drop(['AS_OF_PERIOD','CREATED_BY','CREATED_ON','PERIOD_ID','STYLE','TXN_CURRENCY','UPDATED_BY','UPDATED_ON',"ULT_PARENT_CD","REGION","ND_IND"],axis=1)
#    preproc.rename(columns={"﻿AS_OF_DATE":u"AS_OF_DATE"},inplace=True)
#    preproc = preproc.pivot(index='id',columns='AS_OF_DATE',values='TOTAL_OUTFLOW')
#    preproc.columns=[datetime.datetime.strptime(x,'%m/%d/%Y %H:%M:%S AM')-datetime.timedelta(days=0.5) for x in preproc.columns]
#    preproc = preproc.transpose().sort_index().transpose()
#    preproc.to_excel(filename_preproc)
#    preproc = []
#
#preproc = pd.ExcelFile(filename_preproc).parse("Sheet1")
##preproc.head()
#if 'id' in preproc.columns:
#    preproc = preproc.set_index('id',drop=True)
##preproc.columns=[datetime.datetime.strptime(x,'%m/%d/%Y %H:%M:%S AM') for x in preproc.columns]
##preproc.head()
#update_end = datetime.datetime.today().strftime('%Y-%m-%d')
#start= [i for i,x in enumerate(date_mapping.index==update_start_trx) if x][0]
#end = [i for i,x in enumerate(date_mapping.index==update_end) if x][0]
#   
#table_name = "TBL_DEPOSIT_TRX_UP_GTRM"
#for i in range(start,end+1):
##    i = 343
##   i = 340
#    if not any(date_mapping.index[i]==preproc.columns):
#        date = date_mapping.index[i].strftime("%Y%m%d")
#        table_tmp = read_db(table_name,date)
#        if len(table_tmp) >0:
#            print(date)
#            print(table_tmp['TRANSACTION_TYPE'].unique())
#            table_tmp = table_tmp[table_tmp['TRANSACTION_TYPE']=='C']
#            date = table_tmp['AS_OF_DATE'].drop_duplicates().values
#            table_tmp.reset_index(inplace=True)  
#            table_tmp.index = table_tmp["ULT_PARENT_CD"]+table_tmp["REGION"]+table_tmp["ND_IND"]
#            table_tmp = table_tmp.drop(["ULT_PARENT_CD","REGION","ND_IND",'AMOUNT','TRANSACTION_TYPE','AS_OF_DATE'],axis=1)
#            table_tmp.columns = date
#            preproc = preproc.join(table_tmp,how="outer")
#            table_tmp = []
#            gc.collect()
#preproc.to_excel(filename_preproc)
#preproc = []

###### convert excels to hdf db
imp.reload(UpdateDatabase_DB)

##### Daily Balance data
def update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type,update_start=update_start,attribute_groupby='ULT_PARENT_CD',table_name = "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM"):
    if not os.path.exists(file_total_deposit):
        Balance_History = pd.DataFrame()
        Balance_History_client = pd.DataFrame()
        for i in range(len(date_mapping)):
            date = date_mapping.index[i].strftime("%Y%m%d")
#            table_name = "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM"
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
                    
                balance = summarizeTable(table_tmp,attribute_groupby,balance_type,fillna=True)
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
        Balance_History_client = Balance_History_client.set_index([attribute_groupby])
    
        update_end = datetime.datetime.today().strftime('%Y-%m-%d')
        start= [i for i,x in enumerate(date_mapping.index==update_start) if x][0]
        end = [i for i,x in enumerate(date_mapping.index==update_end) if x][0]
           
        for i in range(start,end+1):
            if not any(date_mapping.index[i]==Balance_History.columns):
                date = date_mapping.index[i].strftime("%Y%m%d")
#                table_name = "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM"
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
#                table_name = "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM"
                table_tmp = read_db(table_name,date)
                table_tmp = table_tmp[~pd.isnull(table_tmp.index)]
                if len(table_tmp) >0:
                    balance = summarizeTable(table_tmp,attribute_groupby,balance_type,fillna=True)
                    date = table_tmp['AS_OF_DATE'].drop_duplicates().values
                    balance.columns = date
                    Balance_History_client = Balance_History_client.join(balance,how="outer")
                    table_tmp = []
                    gc.collect()
    
        Balance_History[~pd.isnull(Balance_History.index)].to_excel(file_total_deposit)
        Balance_History_client[~pd.isnull(Balance_History_client.index)].to_excel(file_total_deposit_client)
            #copy to alternative folder
    if COPY_TO_FOLDER:
        Balance_History[~pd.isnull(Balance_History.index)].to_excel(file_total_deposit_client.replace(misc_dir,dir_excel_copyto))
        Balance_History_client[~pd.isnull(Balance_History_client.index)].to_excel(file_total_deposit_client.replace(misc_dir,dir_excel_copyto))

#balance_type = "TOTAL_OUTFLOW"
#file_total_deposit = misc_dir+"/RawData/PaymentTableByParentRegion.xlsx"
#file_total_deposit_client = misc_dir+"/RawData/PaymentTableByParent.xlsx"
#update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type,table_name="PBUP_OPSDEP_PREPROC")
#if UPDATE_RAI_GTRM:
#    file_total_deposit = misc_dir_rai_gtrm+"/RawData/TotalDepositTableByParentRegion.xlsx"
#    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/TotalDepositTableByParent.xlsx"
#    update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type,table_name = table_name_up_rai_gtrm)

balance_type = "DAILY_SPOT_BAL_USD"
file_total_deposit = misc_dir+"/RawData/TotalDepositTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/TotalDepositTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)
if UPDATE_RAI_GTRM:
    file_total_deposit = misc_dir_rai_gtrm+"/RawData/TotalDepositTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/TotalDepositTableByParent.xlsx"
    update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type,table_name = table_name_up_rai_gtrm)

balance_type = "OPERATIONAL_BAL_USD"
file_total_deposit = misc_dir+"/RawData/OperationalDepositTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/OperationalDepositTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)
if UPDATE_RAI_GTRM:
    file_total_deposit = misc_dir_rai_gtrm+"/RawData/OperationalDepositTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/OperationalDepositTableByParent.xlsx"
    update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type,table_name = table_name_up_rai_gtrm)

balance_type = "binding_constraint"
file_total_deposit = misc_dir+"/RawData/BindingConstraintTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/BindingConstraintTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)
if UPDATE_RAI_GTRM:
    file_total_deposit = misc_dir_rai_gtrm+"/RawData/BindingConstraintTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/BindingConstraintTableByParent.xlsx"
    update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type,table_name = table_name_up_rai_gtrm)

balance_type = "BEHAVIORAL_GROUP"
file_total_deposit = misc_dir+"/RawData/BehaviorGroupTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/BehaviorGroupTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)
if UPDATE_RAI_GTRM:
    file_total_deposit = misc_dir_rai_gtrm+"/RawData/BehaviorGroupTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/BehaviorGroupTableByParent.xlsx"
    update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type,table_name = table_name_up_rai_gtrm)

balance_type = "COHORT_RATIO"
file_total_deposit = misc_dir+"/RawData/CohortRatioTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/CohortRatioTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)
if UPDATE_RAI_GTRM:
    file_total_deposit = misc_dir_rai_gtrm+"/RawData/CohortRatioTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/CohortRatioTableByParent.xlsx"
    update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type,table_name = table_name_up_rai_gtrm)

balance_type = "AVG_PAYMENT"
file_total_deposit = misc_dir+"/RawData/AveragePaymentTableByParentRegion.xlsx"
file_total_deposit_client = misc_dir+"/RawData/AveragePaymentTableByParent.xlsx"
update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type)
if UPDATE_RAI_GTRM:
    file_total_deposit = misc_dir_rai_gtrm+"/RawData/AveragePaymentTableByParentRegion.xlsx"
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/AveragePaymentTableByParent.xlsx"
    update_daily_balance_master_file(file_total_deposit,file_total_deposit_client,date_mapping,balance_type,table_name = table_name_up_rai_gtrm)

def update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,update_start=update_start,table_name = "OPS_DEPOSIT_PBUP_FBO_DDA_GTRM"):
    if not os.path.exists(file_total_deposit_client):
        Balance_History_client = pd.DataFrame()
        for i in range(len(date_mapping)):
#        for i in [247,248]:
#       i = 331
            date = date_mapping.index[i].strftime("%Y%m%d")
#            table_name = "OPS_DEPOSIT_PBUP_FBO_DDA_GTRM"
            table_tmp = read_db(table_name,date)

            if len(table_tmp) >0:
                table_tmp = process_dda_GTRM(table_tmp)
                balance = summarizeTable(table_tmp,attribute_groupby,balance_type,fillna=True)
                
                if len(Balance_History_client) == 0:
                    Balance_History_client = balance
                else:
                    Balance_History_client = Balance_History_client.join(balance,how="outer")
                table_tmp = []
                gc.collect()
    else:
        Balance_History_client = pd.ExcelFile(file_total_deposit_client).parse("Sheet1")
        Balance_History_client = Balance_History_client.set_index([attribute_groupby])
    
        update_end = datetime.datetime.today().strftime('%Y-%m-%d')
        start= [i for i,x in enumerate(date_mapping.index==update_start) if x][0]
        end = [i for i,x in enumerate(date_mapping.index==update_end) if x][0]
           
        for i in range(start,end+1):
            if not any(date_mapping.index[i]==Balance_History_client.columns):
                date = date_mapping.index[i].strftime("%Y%m%d")
#                table_name = "OPS_DEPOSIT_PBUP_FBO_DDA_GTRM"
                table_tmp = read_db(table_name,date)
                if len(table_tmp) >0:
                    table_tmp = process_dda_GTRM(table_tmp)
                    balance = summarizeTable(table_tmp,attribute_groupby,balance_type,fillna=True)
                    date = table_tmp['AS_OF_DATE'].drop_duplicates().values

                    Balance_History_client = Balance_History_client.join(balance,how="outer")
                    table_tmp = []
                    gc.collect()
    
    Balance_History_client[~pd.isnull(Balance_History_client.index)].to_excel(file_total_deposit_client)
    #copy to alternative folder
    if COPY_TO_FOLDER:
        Balance_History_client[~pd.isnull(Balance_History_client.index)].to_excel(file_total_deposit_client.replace(misc_dir,dir_excel_copyto))

#parent and spot
balance_type = column_name_spot
attribute_groupby = "FUND_NUM"
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
#parent and opetaional
balance_type = column_name_operational
attribute_groupby = "FUND_NUM"
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

#parent and spot
balance_type = column_name_spot
attribute_groupby = column_name_parent
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
#parent and opetaional
balance_type = column_name_operational
attribute_groupby = column_name_parent
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

#parentregion and spot
balance_type = column_name_spot
attribute_groupby = "ID"
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
#parentregion and operational
balance_type = column_name_operational
attribute_groupby = "ID"
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

#behavior group and spot
balance_type = column_name_spot
attribute_groupby = column_name_group
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
#behavior group and operational
balance_type = column_name_operational
attribute_groupby = column_name_group
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

#region and spot
balance_type = column_name_spot
attribute_groupby = column_name_region
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
#region and operational
balance_type = column_name_operational
attribute_groupby = column_name_region
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

#legal entity and spot
balance_type = column_name_spot
attribute_groupby = column_name_company
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
#legal entity and operational
balance_type = column_name_operational
attribute_groupby = column_name_company
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

#currency and spot
balance_type = column_name_spot
attribute_groupby = column_name_currency
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
#currency and operational
balance_type = column_name_operational
attribute_groupby = column_name_currency
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

#product type and spot
balance_type = column_name_spot
attribute_groupby = 'ProductTypes'
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
#product type and operational
balance_type = column_name_operational
attribute_groupby = 'ProductTypes'
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

#country of domicile and spot
balance_type = column_name_spot
attribute_groupby = 'COUNTRY_DOMICILE'
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
#country of domicile and operational
balance_type = column_name_operational
attribute_groupby = 'COUNTRY_DOMICILE'
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

#investment style and spot
balance_type = column_name_spot
attribute_groupby = column_name_stylegroup
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
#investment style and operational
balance_type = column_name_operational
attribute_groupby = column_name_stylegroup
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

#business division and spot
balance_type = column_name_spot
attribute_groupby = column_name_division
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
#business division and operational
balance_type = column_name_operational
attribute_groupby = column_name_division
file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
if UPDATE_RAI_GTRM:
    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

##business business lines
#balance_type = column_name_spot
#attribute_groupby = column_name_bu
#file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
#update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
##if UPDATE_RAI_GTRM:
##    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
##    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)
##business division and operational
#balance_type = column_name_operational
#attribute_groupby = column_name_bu
#file_total_deposit_client = misc_dir+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
#update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby)
##if UPDATE_RAI_GTRM:
##    file_total_deposit_client = misc_dir_rai_gtrm+"/RawData/"+balance_type+"_By_"+ attribute_groupby+".xlsx"
##    update_daily_balance_master_file_from_DDA(file_total_deposit_client,date_mapping,balance_type,attribute_groupby,table_name = table_name_dda_rai_gtrm)

end_time = time.time()
print("Common Table Preprocess Time is "+str(round(end_time-start_time,0))+" seconds")
