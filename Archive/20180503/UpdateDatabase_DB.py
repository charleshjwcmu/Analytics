# -*- coding: utf-8 -*-
"""
convert data and save to online shared folder

@author: e620927
"""
PATH_DB = "Z:/FTDRDataBase/"
PRODUCTION_ENVRIONMENT = True
################# Global Parameters  #################
if PRODUCTION_ENVRIONMENT:
    days_need_update = 10
else:
    days_need_update = 30
#    days_need_update = 273
    
import os
import pandas as pd
import datetime
import time
import imp

#tmp_cwd = os.getcwd()
#os.chdir("Z:/Charles/PyLibrary")
import sys
sys.path.insert(0, "Z:/Charles/PyLibrary")
import Library_ReadData
imp.reload(Library_ReadData)
from Library_ReadData import read_database, delete_database
#os.chdir(tmp_cwd)

UPDATE = True
#UPDATE = False

#excel format table
def update_db_excel(table_name,raw_directory,out_directory,date_list=None):
#    date_list=["20170930"]
    if date_list == None:
        base = datetime.datetime.today()
        date_list = [(base - datetime.timedelta(days=x)).strftime("%Y%m%d") for x in range(0, days_need_update)]

    if not os.path.exists(out_directory):
        os.makedirs(out_directory)
                 
    for date in date_list:
    #    date = date_list[276]
        file_name_source = raw_directory+"/"+table_name+"_"+date+".xlsx"
        file_name_output = out_directory+"/"+table_name+"_"+date+".hdf"
        
        if os.path.exists(file_name_source) and (not os.path.exists(file_name_output)):
            print(table_name+"/"+"New Source File Found: "+file_name_source)
            start_time = time.time()
            data = pd.ExcelFile(file_name_source)
            
            data_excel = data.parse("Sheet1")
            
            if "Sheet2" in data.sheet_names:
                data_excel2 = data.parse("Sheet2")
                if (len(data_excel2.index) > 0):
                    data_excel = data_excel.append(data_excel2)
            
            print(table_name+"/"+"Read Excel data using " + str(round(time.time()-start_time,0)) + " seconds")
#            print(data_excel["AS_OF_DATE"].unique())
            start_time = time.time()
            data_excel.to_hdf(file_name_output,"w",table=True)
            print(table_name+"/"+"Convert HDF data using " + str(round(time.time()-start_time,0)) + " seconds")
            
        elif not os.path.exists(file_name_source):
            print(table_name+"/"+date + ": Source File Not Found")
        elif os.path.exists(file_name_output):
            print(table_name+"/"+date + ": Data Conversion has already completed")
if UPDATE:
    table_names = ["TRANS_PROD_TREASURY_COMMON_DEPOSITS","OPSDEP_PREPROC",\
                   "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP","PBUP_OPSDEP_FUNDBAL_OUTPUT_DDA", "OPS_DEPOSIT_PBUP_FBO_DDA_VW","OPS_DEPOSIT_PBUP_FBO_DDA_GTRM","PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM",\
                   "PBUP_OPSDEP_RAI_OUTPUT","REF_OPSDEP_CATEGORY", "TBL_REF_HEDGE_FUND","REF_CDMS_CUSTOMER","OPSDEP_MASTERATTR","TBL_REF_PRODUCT_LOOKUP"]
    for table_name in table_names:
        raw_directory = PATH_DB+table_name+"/RawData"
        out_directory = PATH_DB+table_name+"/HdfData"
        update_db_excel(table_name,raw_directory,out_directory)

#table_names = ["TBL_REF_PRODUCT_LOOKUP"]
#for table_name in table_names:
#    raw_directory = PATH_DB+table_name+"/RawData"
#    out_directory = PATH_DB+table_name+"/HdfData"
#    update_db(table_name,raw_directory,out_directory,date_list=["20151123"])

def update_db_txt(table_name,raw_directory,out_directory,date_list=None):
#    date_list=["20171108"]
    if date_list == None:
        base = datetime.datetime.today()
        date_list = [(base - datetime.timedelta(days=x)).strftime("%Y%m%d") for x in range(0, days_need_update)]

    if not os.path.exists(out_directory):
        os.makedirs(out_directory)
    
    for date in date_list:
    #    date = date_list[0]
        file_name_source = raw_directory+"/"+table_name+"_"+date+".txt"
        file_name_output = out_directory+"/"+table_name+"_"+date+".hdf"
        
        if os.path.exists(file_name_source) and (not os.path.exists(file_name_output)):
            print(table_name+"/"+"New Source File Found: "+file_name_source)
            start_time = time.time()
            
            data = pd.read_table(file_name_source,low_memory=False)
            print(table_name+"/"+"Read Excel data using " + str(round(time.time()-start_time,0)) + " seconds")
#            print(data_excel["AS_OF_DATE"].unique())
    
            start_time = time.time()
            data.to_hdf(file_name_output,"w",table=True)
            print("Convert HDF data using " + str(round(time.time()-start_time,0)) + " seconds")
            
        elif not os.path.exists(file_name_source):
            print(table_name+"/"+date + ": Source File Not Found")
        elif os.path.exists(file_name_output):
            print(table_name+"/"+date + ": Data Conversion has already completed")

if UPDATE:
    table_names = ["PBUP_OPSDEP_DDA_XREF","PBUP_OPSDEP_VOL_MITG_SPOT_BAL","PBUP_OPSDEP_PREPROC","TBL_DEPOSIT_TRX"]
    for table_name in table_names:
        raw_directory = PATH_DB+table_name+"/RawData"
        out_directory = PATH_DB+table_name+"/HdfData"
        update_db_txt(table_name,raw_directory,out_directory)

# Read data
def read_db(table_name, date_list=None):
#    date_list = "20171026"
#    date_list = formatDate(dates[table_name])
    if date_list == None:
        base = datetime.datetime.today()
        date_list = [(base - datetime.timedelta(days=x)).strftime("%Y%m%d") for x in range(0, 10)]
    elif date_list.__class__ is str:
        date_list = [date_list]

    if table_name == "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP":
        # table PBUP_OPSDEP_FUNDBAL_OUTPUT_UP
        out_directory = PATH_DB+table_name+"/HdfData/"
        file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
        
        required_attributes = ["AS_OF_DATE","BEHAVIORAL_GROUP","COHORT_RATIO","ND_IND","REGION","ULT_PARENT_CD",\
        "DAILY_SPOT_BAL_USD","EXCESS_BAL_USD","OPERATIONAL_BAL_USD","SPOT_BAL_PERC","SPOT_BAL_PERC_DISCRE"]
        data_all = read_database(file_list,required_attributes)
    elif table_name == "PBUP_OPSDEP_FUNDBAL_OUTPUT_DDA":
        # table PBUP_OPSDEP_FUNDBAL_OUTPUT_UP
        out_directory = PATH_DB+table_name+"/HdfData/"
        file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
        
        required_attributes = ["AS_OF_DATE", "DDA", "FUND_NUM", "DAILY_SPOT_BAL_USD","EXCESS_BAL_USD","OPERATIONAL_BAL_USD","SPOT_BAL_PERC"]
        data_all = read_database(file_list,required_attributes)        
    elif table_name == "OPS_DEPOSIT_PBUP_FBO_DDA_VW":
        # table OPS_DEPOSIT_PBUP_FBO_DDA_VW
        out_directory = PATH_DB+table_name+"/HdfData/"
        file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
        
        required_attributes = ["AS_OF_DATE","PERIOD_ID","COMPANY_DESC","TRANSACTION_CURRENCY",\
        "BU_LVL3_DESC","BU_LVL4_DESC","BU_LVL5_DESC","BU_LVL6_DESC",\
        "CDMS_ULT_PRNT_CODE","REGION","ND_IND","CDMS_CUST_CODE","CDMS_CUST_NAME","FUND_NUM","ULT_PARENT_NAME","SOURCE_CONTRACT_NUM","SOURCE_SYSTEM_ID",\
        "DAILY_EXCESS_BAL_USD","DAILY_OPERATIONAL_BAL_USD","DAILY_SPOT_BAL_USD",\
        "OPERATIONAL_BINDING_CONSTRAINT","SPOT_BAL_PERC","WGT_PRINCIPAL_BAL_USD"]
        data_all = read_database(file_list,required_attributes)
    elif table_name == "PBUP_OPSDEP_RAI_OUTPUT":
        out_directory = PATH_DB+table_name+"/HdfData/"
        file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
        required_attributes = ["CREATED_ON","PERIOD_ID",\
         "AVG_PAYMENT","BEHAVIORAL_GROUP","ND_IND","ULT_PARENT_CD","REGION",\
         "COHORT_RATIO","PERC_BAL_USD_75","PERC_BAL_USD_95","UPPER_BOUND_1","UPPER_BOUND_2"] 
        data_all = read_database(file_list,required_attributes)
    elif table_name == "PBUP_OPSDEP_RAI_GTRM":
        out_directory = PATH_DB+table_name+"/HdfData/"
        file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
        required_attributes = ["AVG_PAYMENT","BEHAVIORAL_GROUP","ND_IND","ULT_PARENT_CD","REGION",\
         "COHORT_RATIO","PERC_BAL_USD_75","PERC_BAL_USD_95","UPPER_BOUND_1","UPPER_BOUND_2"] 
        data_all = read_database(file_list,required_attributes)
        
    elif table_name == "TRANS_PROD_TREASURY_COMMON_DEPOSITS":
        out_directory = PATH_DB+table_name+"/HdfData/"
        file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
        required_attributes = ["AS_OF_DATE","ASSET_LIABILITY_CODE","TDR_STATUS_FLAG",\
        "COUNTRY_DOMICILE","FUND_NUM","INTERNAL_COMPANY_FLAG",\
        "PRIN_PARENT_NODE","PRIN_PARENT_NODE_DESC","PRINCIPAL_BAL_USD",\
        "PRODUCT_AGGREGATE_DESC","PRODUCT_GROUP_DESC",\
        "IFS_PRODUCT","IFS_COMPANY","PROGRESSION_RESP_CENTER","PRODUCT_SUB_TYPE_DESC","PRODUCT_TYPE_DESC","SOURCE_CONTRACT_NUM",\
        "SOURCE_PRODUCT_TYPE","TRANSACTION_CURRENCY"]
#        "INTEREST_RATE","SITE","SOURCE_SYSTEM_ID","PERIOD_ID","BRANCH","CDMS_COUNTERPARTY_ID","CDMS_COUNTERPARTY_NAME",
        
        data_all = read_database(file_list,required_attributes)
    elif table_name == "PBUP_OPSDEP_DDA_XREF":
        out_directory = PATH_DB+table_name+"/HdfData/"
        file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
        required_attributes = ["CREATED_ON","DDA","FUND_NUM",\
        "REGION","ULT_PARENT_CD","UPDATED_ON"]
        data_all = read_database(file_list,required_attributes)
    elif table_name == "REF_OPSDEP_CATEGORY":
        out_directory = PATH_DB+table_name+"/HdfData/"
        file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
        required_attributes = ["CREATED_ON","SOURCE_CONTRACT_NUM","UPDATED_ON"]
        data_all = read_database(file_list,required_attributes)
    elif table_name == "TBL_REF_HEDGE_FUND":
        out_directory = PATH_DB+table_name+"/HdfData/"
        file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
        required_attributes = ["CREATED_ON","FUND_NUMBER","UPDATED_ON"]
        data_all = read_database(file_list,required_attributes)
    elif table_name == "TBL_REF_PRODUCT_LOOKUP":
        out_directory = PATH_DB+table_name+"/HdfData/"
        file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
        required_attributes = ["ACCT_HIER_SEQ_CODE","EXCLUDES","IFS_PRODUCT","PROD_GROUP"]
        data_all = read_database(file_list,required_attributes)
    elif table_name in ["REF_CDMS_CUSTOMER", "OPSDEP_MASTERATTR","PBUP_OPSDEP_VOL_MITG_SPOT_BAL","OPSDEP_PREPROC","PBUP_OPSDEP_PREPROC","TBL_DEPOSIT_TRX","TBL_DEPOSIT_TRX_UP_GTRM",\
                        "PBUP_OPSDEP_FUNDBAL_OUTPUT_UP_GTRM","OPS_DEPOSIT_PBUP_FBO_DDA_GTRM","OPS_DEPOSIT_DDA_RAI_GTRM","OPS_DEPOSIT_UP_RAI_GTRM"]:
        out_directory = PATH_DB+table_name+"/HdfData/"
        file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
        data_all = read_database(file_list)
    elif table_name in ["TotalDepositTableByParent","TotalDepositTableByParentRegion",\
    "OperationalDepositTableByParent","OperationalDepositTableByParentRegion",\
    "BindingConstraintTableByParent","BindingConstraintTableByParentRegion",\
    "BehaviorGroupTableByParent","BehaviorGroupTableByParentRegion",\
    "CohortRatioTableByParent","CohortRatioTableByParentRegion",\
    "AveragePaymentTableByParent","AveragePaymentTableByParentRegion",\
    "Total_Deposits_By_BehaviorGroup","Operational_Deposits_By_BehaviorGroup",\
    "Total_Deposits_By_Region","Operational_Deposits_By_Region",\
    "Total_Deposits_By_LegalEntity","Operational_Deposits_By_LegalEntity",\
    "Total_Deposits_By_Currency","Operational_Deposits_By_Currency",\
    "Total_Deposits_By_ProductTypes","Operational_Deposits_By_ProductTypes",\
    "Total_Deposits_By_FundStyle","Operational_Deposits_By_FundStyle",\
    "Total_Deposits_By_COUNTRY_DOMICILE","Operational_Deposits_By_COUNTRY_DOMICILE",\
    "Total_Deposits_By_BusinessUnit","Operational_Deposits_By_BusinessUnit",\
    "Total_Deposits_By_BusinessLines","Operational_Deposits_By_BusinessLines",\
    "Total_Deposits_By_ParentCompy","Operational_Deposits_By_ParentCompy",\
    "Total_Deposits_By_FUND_NUM","Operational_Deposits_By_FUND_NUM"]:
        out_directory = PATH_DB+"MISC_TABLES/RawData/"
        file = out_directory + table_name + ".xlsx"
        data_all = pd.ExcelFile(file).parse("Sheet1",index_col=0)
        data_all = data_all.transpose().sort_index().transpose()
    elif table_name == "Mapping_RESPCenter_BU":
        out_directory = PATH_DB+"MISC_TABLES/RawData/"
        file = out_directory + table_name + ".xlsx"
        data_all = pd.ExcelFile(file).parse("Sheet1",index_col=0)
    else:
        print("Cannot find the database. Check table name " + table_name)
        return()
    
    return(data_all)
    
def delete_db(table_name,date_list):
    if date_list.__class__ is str:
        date_list = [date_list]
    out_directory = PATH_DB+table_name+"/HdfData/"
    file_list = [out_directory + table_name + "_" + date + ".hdf" for date in date_list]
    delete_database(file_list)

    out_directory = PATH_DB+table_name+"/RawData/"
    file_list = [out_directory + table_name + "_" + date + ".xlsx" for date in date_list]
    delete_database(file_list)
