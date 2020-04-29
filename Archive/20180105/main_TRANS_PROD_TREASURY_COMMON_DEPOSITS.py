# -*- coding: utf-8 -*-
"""
Created on Wed Nov  1 16:01:25 2017
Analyze treasury common deposits file
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
    daily_operational_balance_input_file = "Z:/Charles/ORM/Data/DAILY_OPERATIONAL_BALANCE_SUMMARY_20171129.xlsx"
    daily_operational_balance_gtrm_input_file = "Z:/Charles/ORM/Data/DAILY_OPERATIONAL_BALANCE_SUMMARY_GTRM.xlsx"
    num_top_funds_to_plot = 40
    TEST_ENVIRONMENT = False
else:
    code_dir = "Z:/Charles/ORM/SourceCodes"
    output_dir = "Z:/Charles/ORM/Output_Test"
    pdf_dir = "Z:/Charles/ORM/PDFReport_Test"
    mycommontable_dir = "Z:/FTDRDataBase/OPS_DEPOSIT_PBUP_FBO_DDA_GTRM"
    input_file = "Z:/Charles/ORM/Data/Parameters.xlsx"
    daily_operational_balance_input_file = "Z:/Charles/ORM/Data/DAILY_OPERATIONAL_BALANCE_SUMMARY_20171129.xlsx"
    daily_operational_balance_gtrm_input_file = "Z:/Charles/ORM/Data/DAILY_OPERATIONAL_BALANCE_SUMMARY_GTRM.xlsx"
    num_top_funds_to_plot = 40
    TEST_ENVIRONMENT = True

################# Load Libraries #################
import os
import pandas as pd
import time
import numpy as np
import matplotlib.pyplot as plt
#from pandas.tools.plotting import table
from matplotlib.ticker import FuncFormatter
import datetime
import dateutil
import subprocess
#import shutil
import imp
from scipy import stats
import gc

# Load local functions
os.chdir(code_dir)
import UpdateDatabase_DB
imp.reload(UpdateDatabase_DB)
from UpdateDatabase_DB import read_db
import Functions_Analysis
imp.reload(Functions_Analysis)
from Functions_Analysis import summarizeTableByAttribute, summarizeTable#, summarizeTableRatioOfTwoFactor, listisin
#################### Configuration  ####################
#Setup parameters
Parameters=pd.ExcelFile(input_file).parse("Sheet1")
benchmarkdate = Parameters.ix["BenchmarkDate"][0] 

if TEST_ENVIRONMENT:
    start_date = datetime.datetime.strptime('2017-12-1 00:00:00', '%Y-%m-%d %H:%M:%S')   # historical data start date - before benchmark date
    benchmarkdate = start_date
    end_date = datetime.datetime.strptime('2017-12-13 00:00:00', '%Y-%m-%d %H:%M:%S')     # historical data end date - later than benchmark date
else:
    benchmarkdate = [datetime.datetime.strptime('2017-11-07 00:00:00', '%Y-%m-%d %H:%M:%S'),\
    datetime.datetime.strptime('2017-11-08 00:00:00', '%Y-%m-%d %H:%M:%S'),\
    datetime.datetime.strptime('2017-11-09 00:00:00', '%Y-%m-%d %H:%M:%S'),\
    datetime.datetime.strptime('2017-11-10 00:00:00', '%Y-%m-%d %H:%M:%S'),\
    datetime.datetime.strptime('2017-11-13 00:00:00', '%Y-%m-%d %H:%M:%S'),\
    datetime.datetime.strptime('2017-11-14 00:00:00', '%Y-%m-%d %H:%M:%S')]

    if not pd.isnull(Parameters.ix['HistoryEndDate'][0]):
        end_date = Parameters.ix["HistoryEndDate"][0]
    else:
        end_date = datetime.datetime.today()
        
    if not pd.isnull(Parameters.ix['HistoryStartDate'][0]):
        start_date = Parameters.ix["HistoryStartDate"][0]
    else:
        start_date = end_date - datetime.timedelta(days=33)
#        if start_date>min(benchmarkdate):
#            start_date = min(benchmarkdate)

timewindow = [(end_date - datetime.timedelta(days=x)).strftime("%Y%m%d") for x in range((end_date-start_date).days+1)]
benchmarkwindow = [x.strftime("%Y%m%d") for x in benchmarkdate]
timewindow = sorted(list(set(benchmarkwindow+timewindow)))

date_newmodel = datetime.datetime.strptime("2017-4-26 00:00:00", '%Y-%m-%d %H:%M:%S')
# create folder hierarchy
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
if not os.path.exists(pdf_dir):
    os.makedirs(pdf_dir)
if not os.path.exists(mycommontable_dir):
    os.makedirs(mycommontable_dir)
if not os.path.exists(mycommontable_dir+"/RawData"):
    os.makedirs(mycommontable_dir+"/RawData")
    
day = datetime.date.today().strftime("%Y%m%d")
plot_folder = output_dir+"/Output_ORM_"+day
if not os.path.exists(plot_folder):
    os.mkdir(plot_folder)

# formats and styles
pd.options.mode.chained_assignment = None
plt.style.use('ggplot')
styles = ['b-','g-','r-','c-','m-','y-','k-']
styles2 = ['b--','g--','r--','c--','m--','y--','k--']
#styles = ['bs-','go-','r^-','cs-','mo-','y^-','kp-']
#styles2 = ['bs--','go--','r^--','cs--','mo--','y^--','kp--']

def percentages(x):
    return '{:.1%}'.format(x)
def millions(x, pos=0):
    return '{:,.1f}'.format(x*1e-6)
def billions(x, pos=0):
    return '{:,.1f}B'.format(x*1e-9)
def billionsinteger(x, pos=0):
    return '{:,.0f}B'.format(x*1e-9)

formatter_millions = FuncFormatter(millions)
formatter_billions = FuncFormatter(billions)
formatter_billionsinteger = FuncFormatter(billionsinteger)

#################### Finish Configuration ####################

start_time = time.time()
#################### Read and Processing Data #########
table_name = "OPS_DEPOSIT_PBUP_FBO_DDA_GTRM"
data_all = read_db(table_name,timewindow)
data_all.index=data_all['ULT_PARENT_CD']+data_all['REGION']+data_all['ND_IND']
data_all['AS_OF_DATE'] = [datetime.datetime.strptime(x, '%Y-%m-%d %H:%M:%S') for x in data_all['AS_OF_DATE']]

data_all['ID'] = data_all.index

#rename columns
column_name_operational_orig = "DAILY_OPERATIONAL_BAL_USD"
column_name_spot_orig = "DAILY_SPOT_BAL_USD"
column_name_excess_orig = "DAILY_EXCESS_BAL_USD"

column_name_company_orig = "IFS_COMPANY"
column_name_bu_orig = "BU_LVL3_DESC"
column_name_region_orig = "REGION"
column_name_group_orig = "BEHAVIORAL_GROUP"
column_name_currency_orig = "TRANSACTION_CURRENCY"
column_name_parent_orig = "ULT_PARENT_CD"
column_name_parentname_orig = 'ULT_PARENT_NAME'
column_name_style_orig = "STYLE"
column_name_stylegroup_orig = "STYLE_GROUP"
column_name_division_orig = "DIVISION_GROUP"
column_name_market_orig = "GROUP_MARKET_SEGMENT"

column_name_productagg_type_orig = "PRODUCT_AGGREGATE_DESC"
column_name_product_type_orig = 'PRODUCT_TYPE_DESC'
column_name_product_subtype_orig = "PRODUCT_SUB_TYPE_DESC"
column_name_GL_type_orig = "PRIN_PARENT_NODE_DESC"

column_name_operational = "Operational_Deposits"
column_name_spot = "Total_Deposits"
column_name_excess = "Excess_Deposits"

column_name_company = "LegelEntity"
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

#rename
data_all.rename(columns=
                {column_name_spot_orig:column_name_spot,\
                 column_name_operational_orig:column_name_operational,\
                 column_name_excess_orig:column_name_excess,\
                 column_name_company_orig:column_name_company,\
                 column_name_bu_orig:column_name_bu,\
                 column_name_region_orig:column_name_region,\
                 column_name_group_orig:column_name_group,\
                 column_name_currency_orig:column_name_currency,\
                 column_name_parent_orig:column_name_parent,\
                 column_name_parentname_orig:column_name_parentname,\
                 column_name_style_orig:column_name_style,\
                 column_name_stylegroup_orig:column_name_stylegroup,\
                 column_name_division_orig:column_name_division,\
                 column_name_market_orig:column_name_market,\
                 column_name_product_type_orig:column_name_product_type,\
                 column_name_productagg_type_orig:column_name_productagg_type,\
                 column_name_product_subtype_orig:column_name_product_subtype,\
                 column_name_GL_type_orig:column_name_GL_type\
                 },inplace=True)

attribute_aggs = [column_name_spot,column_name_operational,column_name_excess]

select_null = pd.isnull(data_all[column_name_operational])
if not all(abs(data_all.loc[~select_null,column_name_spot]-data_all.loc[~select_null,column_name_operational]-data_all.loc[~select_null,column_name_excess])<1):
    print("***ERROR: spot != operational + excess")
data_all.loc[select_null,column_name_excess]=data_all.loc[select_null,column_name_spot]

data_all.loc[data_all[column_name_company]==2001,column_name_company]="SSBT USA"
data_all.loc[data_all[column_name_company]==2002,column_name_company]="SSBT Hong Kong"
data_all.loc[data_all[column_name_company]==2004,column_name_company]="SSBT London"
data_all.loc[data_all[column_name_company]==2006,column_name_company]="SSBT Canada"
data_all.loc[data_all[column_name_company]==2011,column_name_company]="SSBT Sydney"
data_all.loc[data_all[column_name_company]==2014,column_name_company]="SSBT Singapore"
data_all.loc[data_all[column_name_company]==2021,column_name_company]="SSBT GmbH Germany"
data_all.loc[data_all[column_name_company]==2022,column_name_company]="SSBT GmbH Luxembourg"
data_all.loc[data_all[column_name_company]==2023,column_name_company]="SSBT GmbH Austria"
data_all.loc[data_all[column_name_company]==2026,column_name_company]="SSBT GmbH Switzerland"
data_all.loc[data_all[column_name_company]==2028,column_name_company]="SSBT GmbH Amsterdam"
data_all.loc[data_all[column_name_company]==2029,column_name_company]="SSBT GmbH Italy"
data_all.loc[data_all[column_name_company]==2613,column_name_company]="SSBT Jersey"
data_all.loc[data_all[column_name_company]==2009,column_name_company]="Others"

prod_type_mapping = {
    "INTEREST BEARING FOREIGN DEPOSITS":"IBDDA",\
    "INTEREST BEARING DOMESTIC DEPOSITS":"IBDDA",\
    "DEMAND DEPOSITS":"DDA",\
    "_Missing":"_Missing"}

data_all[column_name_product_type] = data_all[column_name_product_type].fillna("_Missing")
data_all[column_name_product_subtype] = data_all[column_name_product_subtype].fillna("_Missing")
data_all[column_name_product_type] = [prod_type_mapping[x] if x in prod_type_mapping else x for x in data_all[column_name_product_type]]
data_all[column_name_product_subtype] = [x.replace(" BRANCH","") for x in data_all[column_name_product_subtype]]
data_all["ProductTypes"] = data_all[column_name_product_type]+"-"+data_all[column_name_product_subtype]

#data_all["ProductTypes"].drop_duplicates()==float('nan')

dates=data_all["AS_OF_DATE"].drop_duplicates().sort_values(ascending=False)
today = dates[0]
yesterday = dates[1]
lastweek = today + dateutil.relativedelta.relativedelta(weeks=-1)

if not pd.isnull(Parameters.ix['CurrentDate'][0]):
    currentsnapshot = Parameters.ix["CurrentDate"][0]
else:
    currentsnapshot = max(data_all['AS_OF_DATE'])

#mapping between parent codes with parent names
mapping_code_names = data_all[[column_name_parent,column_name_parentname]].copy()
mapping_code_names = mapping_code_names[~pd.isnull(mapping_code_names.index)]
mapping_code_names = mapping_code_names.drop_duplicates()
mapping_code_names.dropna(subset = [column_name_parentname],inplace = True)
mapping_code_names.drop_duplicates(subset = column_name_parent, keep='first', inplace = True)
#dup_codes = mapping_code_names.index[mapping_code_names.index.duplicated()]
#mapping_code_names.ix[dup_codes]
if not len(mapping_code_names[column_name_parent])==len(mapping_code_names[[column_name_parent]].drop_duplicates()):
    print("Error: The mapping table between parent code and parent name has duplicates")

mapping_code_names = mapping_code_names.set_index([column_name_parent],drop=True)
mapping_code_names[column_name_parentname] = [x.replace("(PARENT)","").replace("CONFIDENTIAL CLIENT","Conf. Client")[:34] for x in mapping_code_names[column_name_parentname]]

#################### Strt Analysis ####################
##### summary of balance trends over time
daily_balance = pd.ExcelFile(daily_operational_balance_input_file).parse('Sheet1')
daily_balance.rename(columns={'SPOT BALANCE':column_name_spot,'OPERATIONAL BALANCE':column_name_operational,'EXCESS BALANCE':column_name_excess},inplace=True)
daily_balance.set_index('AS_OF_DATE',inplace=True)
daily_balance = daily_balance[[column_name_spot,column_name_operational,column_name_excess]]

result = summarizeTableByAttribute(data_all,attribute_aggs)
if os.path.exists(daily_operational_balance_gtrm_input_file):
    daily_balance_gtrm = pd.ExcelFile(daily_operational_balance_gtrm_input_file).parse('Sheet1',index_col=0)
    result_update = result.copy()
    result_update['UPDATED_ON'] = datetime.datetime.today()
    daily_balance_gtrm = daily_balance_gtrm[~daily_balance_gtrm.index.isin(result_update.index)].append(result_update).sort_index()
else:
    daily_balance_gtrm = daily_balance[~daily_balance.index.isin(result.index)].append(result).sort_index()
    daily_balance_gtrm['UPDATED_ON'] = datetime.datetime.today()
daily_balance_gtrm.to_excel(daily_operational_balance_gtrm_input_file)
daily_balance_gtrm = daily_balance_gtrm.dropna()

while True:
    try:
        plt.figure()
        result_figure = result.copy()
        result_figure.columns = [x.replace("_"," ") for x in result_figure.columns]
        figure = result_figure.plot(style=styles,linewidth=1.3)
        figure.yaxis.set_major_formatter(formatter_billionsinteger)
        figure.xaxis.label.set_visible(False)
        plt.title("Trend of Deposit in USD")
        plt.legend(loc=2, bbox_to_anchor=(1,0.8))
        plt.rcParams.update({'font.size': 12})
        fig = figure.get_figure()
        fig.set_size_inches(8, 6)
        fig.savefig(plot_folder+"/Summary_Plot1.png",bbox_inches='tight')
        plt.clf()
        
        plt.figure()
        daily_balance_gtrm_figure = daily_balance_gtrm.copy()
        daily_balance_gtrm_figure.columns = [x.replace("_"," ") for x in daily_balance_gtrm_figure.columns]
        figure = daily_balance_gtrm_figure.plot(style=styles,linewidth=1.3)
        figure.yaxis.set_major_formatter(formatter_billionsinteger)
        figure.xaxis.label.set_visible(False)
        plt.title("Trend of Deposit in USD",fontsize=15)
        plt.legend(loc=2, bbox_to_anchor=(1,0.8))
        plt.rcParams.update({'font.size': 12})
        fig = figure.get_figure()
        fig.set_size_inches(8, 6)
        fig.savefig(plot_folder+"/Summary_DepositChartLong.png",bbox_inches='tight')
        plt.clf()
        break
    except:
        pass
result_tex = result.applymap(billions)
#result.columns = ['DAILY_SPOT_BAL_USD','OPERATIONAL_BAL_USD','EXCESS_BAL_USD']
result_tex = result_tex.sort_index(ascending=False)
result_tex.to_latex(plot_folder+"/Summary_Plot1.tex",longtable=True)
result_tex.to_excel(plot_folder+"/Summary_Plot1.xlsx")

#summary3
def aggregate_summary_table_with_percentiles(summmary_result, attributes = None, reference_dates = None,digit=1,add_total=False):
#    summmary_result = result_total
    summarypage1 = summmary_result.copy()
    if attributes is None:
        attributes = summmary_result.columns
    if reference_dates is None:
        current_date = summarypage1.head(1)
        summarypage1 = summarypage1.head(2).transpose()
        summarypage1.columns = [x.strftime("%Y-%m-%d") for x in summarypage1.columns]
        summarypage1['DailyChg']=summarypage1.iloc[:,0]-summarypage1.iloc[:,1]
    elif reference_dates is not None:
        current_date = summarypage1[summarypage1.index == np.datetime64(reference_dates[0])]
        summarypage1 = summarypage1[summarypage1.index.isin([np.datetime64(x) for x in reference_dates])].transpose()
        summarypage1.columns = [x.strftime("%Y-%m-%d") for x in summarypage1.columns]
    else:
        print("Wrong input.")
        return(None)
    summarypage1 = round(summarypage1*1e-9,digit)
    summarypage1['Average']=round(summmary_result.apply(np.mean,axis=0)*1e-9,digit)
    summarypage1['Max']=round(summmary_result.apply(max,axis=0)*1e-9,digit)
    summarypage1['Min']=round(summmary_result.apply(min,axis=0)*1e-9,digit)
    summarypage1['Std']=round(summmary_result.apply(np.std,axis=0)*1e-9,digit)
    summarypage1 = summarypage1[summarypage1.index.isin(attributes)]
    
    percentiles = []
    for i in range(len(attributes)):
        p=percentages((stats.percentileofscore(summmary_result[attributes[i]], current_date[attributes[i]].values[0], 'rank')/100))
        percentiles.append(p)
    
    summarypage1['Percentile']=percentiles
    summarypage1.index = [x.replace("_"," ") if x is not "_Missing" else x for x in summarypage1.index]
    return(summarypage1)

summmary_result = daily_balance_gtrm.sort_index(ascending=False)[[column_name_spot,column_name_operational,column_name_excess]]
summarypage1 = aggregate_summary_table_with_percentiles(summmary_result)
summarypage1.to_latex(plot_folder+"/Summary_DepositTable.tex",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_DepositTable.xlsx")

summarypage1 = aggregate_summary_table_with_percentiles(summmary_result,reference_dates = [today, yesterday, lastweek])
summarypage1.to_latex(plot_folder+"/Summary_Weekly_DepositTable.tex",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Weekly_DepositTable.xlsx")

summmary_result = daily_balance_gtrm.sort_index(ascending=False)[[column_name_spot,column_name_operational,column_name_excess]]
summmary_result = summmary_result[summmary_result.index>=date_newmodel]
summarypage1 = aggregate_summary_table_with_percentiles(summmary_result)
summarypage1.to_latex(plot_folder+"/Summary_DepositTable_NewModel.tex",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_DepositTable_NewModel.xlsx")

summarypage1 = aggregate_summary_table_with_percentiles(summmary_result,reference_dates = [today, yesterday, lastweek])
summarypage1.to_latex(plot_folder+"/Summary_Weekly_DepositTable_NewModel.tex",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Weekly_DepositTable_NewModel.xlsx")

def aggregate_summary_table_with_percentiles_and_changes(summmary_result, attributes=None,digit=1):
#     summmary_result = result
#    #statistics of level values
#    summarypage1 = summmary_result.copy()
#    current_date = summarypage1.head(1)
#    summarypage1 = summarypage1.head(1).transpose()
#    summarypage1.columns = [x.strftime("%Y-%m-%d") for x in summarypage1.columns]
#    summarypage1 = round(summarypage1*1e-9,digit)
#    summarypage1['Average']=round(summmary_result.apply(np.mean,axis=0)*1e-9,digit)
#    summarypage1['Max']=round(summmary_result.apply(max,axis=0)*1e-9,digit)
#    summarypage1['Min']=round(summmary_result.apply(min,axis=0)*1e-9,digit)
#    summarypage1['Std']=round(summmary_result.apply(np.std,axis=0)*1e-9,digit)
#    summarypage1 = summarypage1[summarypage1.index.isin(attributes)]
#    
#    percentiles = []
#    for i in range(len(attributes)):
#        p=percentages((stats.percentileofscore(summmary_result[attributes[i]], current_date[attributes[i]].values[0], 'rank')/100))
#        percentiles.append(p)
#    summarypage1['Percentile']=percentiles
#    summarypage1.index = [x.replace("_"," ") for x in summarypage1.index]

    # statistics of changes
    if attributes is None:
        attributes=summmary_result.columns
    summarypage2 = summmary_result.copy()
    current_diff = summarypage2.diff(-1).head(1).dropna()
    summarypage2 = current_diff.transpose()
    summarypage2.columns = ["DailyChg"]
    summarypage2 = round(summarypage2*1e-9,digit)
    summarypage2['Average']=round(summmary_result.diff(-1).dropna().apply(np.nanmean,axis=0)*1e-9,digit)
    summarypage2['Max']=round(summmary_result.diff(-1).dropna().apply(np.max,axis=0)*1e-9,digit)
    summarypage2['Min']=round(summmary_result.diff(-1).dropna().apply(np.min,axis=0)*1e-9,digit)
    summarypage2['Std']=round(summmary_result.diff(-1).dropna().apply(np.std,axis=0)*1e-9,digit)
    summarypage2 = summarypage2[summarypage2.index.isin(attributes)]
    percentiles = []
    for i in range(len(attributes)):
#        i = 0
        p=percentages((stats.percentileofscore(summmary_result.diff(-1).dropna()[attributes[i]], current_diff[attributes[i]].values[0], kind = 'rank')/100))
        percentiles.append(p)
    summarypage2['Percentile']=percentiles
    summarypage2.index = [x.replace("_"," ") if x is not "_Missing" else x for x in summarypage2.index]
    
    return(summarypage2)

summmary_result = daily_balance_gtrm.sort_index(ascending=False)[[column_name_spot,column_name_operational,column_name_excess]]
summarypage1 = aggregate_summary_table_with_percentiles_and_changes(summmary_result)
summarypage1.to_latex(plot_folder+"/Summary_DepositChgTable.tex",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_DepositChgTable.xlsx")

summarypage1 = aggregate_summary_table_with_percentiles_and_changes(summmary_result)
summarypage1.to_latex(plot_folder+"/Summary_Weekly_DepositChgTable.tex",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Weekly_DepositChgTable.xlsx")

summmary_result = daily_balance_gtrm.sort_index(ascending=False)[[column_name_spot,column_name_operational,column_name_excess]]
summmary_result = summmary_result[summmary_result.index>=date_newmodel]
summarypage1 = aggregate_summary_table_with_percentiles_and_changes(summmary_result)
summarypage1.to_latex(plot_folder+"/Summary_DepositChgTable_NewModel.tex",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_DepositChgTable_NewModel.xlsx")

summarypage1 = aggregate_summary_table_with_percentiles_and_changes(summmary_result)
summarypage1.to_latex(plot_folder+"/Summary_Weekly_DepositChgTable_NewModel.tex",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Weekly_DepositChgTable_NewModel.xlsx")

daily_balance=[]
daily_balance_gtrm=[]
result_tex=[]
figure=[]
gc.collect()

###Decompose funds
## describe hedge fund excluded
#data_all_hedge_fund = data_all[data_all[column_name_parent]=="UnregulatedFund"]
#result_hedge_fund = summarizeTableByAttribute(data_all_hedge_fund,attribute_aggs)
#result_hedge_fund.rename(columns={column_name_spot:"ExUnregFunds"},inplace=True)
#
##describe funds not existed in RAI files
#data_all_miss_rai = data_all[(pd.isnull(data_all[column_name_group])) & (data_all[column_name_parent]!="UnregulatedFund")]
#result_miss_rai = summarizeTableByAttribute(data_all_miss_rai,attribute_aggs)
#result_miss_rai.rename(columns={column_name_spot:"ExMissRAI"},inplace=True)
#
##describe Non Discretionary Balance
#data_all_YND = data_all[(~pd.isnull(data_all[column_name_group])) & (data_all[column_name_parent]!="UnregulatedFund") & (data_all["ND_IND"]=="Y")]
#data_all_NND = data_all[(~pd.isnull(data_all[column_name_group])) & (data_all[column_name_parent]!="UnregulatedFund") & (data_all["ND_IND"]=="N")]
#result_ynd = summarizeTableByAttribute(data_all_YND,attribute_aggs)
#result_nnd = summarizeTableByAttribute(data_all_NND,attribute_aggs)
#result_ynd.rename(columns={column_name_spot:"NonDiscBal"},inplace=True)
#result_nnd.rename(columns={column_name_spot:"NormalBal"},inplace=True)
#
#result= result.join(result_hedge_fund[["ExUnregFunds"]])
#result= result.join(result_miss_rai[["ExMissRAI"]])
#result["AvailBalUSD"] = result[column_name_spot]-result["ExUnregFunds"]-result["ExMissRAI"]
#result= result.join(result_ynd[["NonDiscBal"]])
#result= result.join(result_nnd[["NormalBal"]])
#
#plt.figure()
#figure = result.plot(style=styles)
#figure.yaxis.set_major_formatter(formatter_billions)
#plt.title("Trend of Deposit (in USD)")
#plt.legend(loc=2, bbox_to_anchor=(1,0.8))
#plt.rcParams.update({'font.size': 12})
#fig = figure.get_figure()
#fig.set_size_inches(8, 6)
#fig.savefig(plot_folder+"/Summary_Plot2.png",bbox_inches='tight')
#plt.clf()
#
#result = result[[column_name_spot,"ExUnregFunds","ExMissRAI","AvailBalUSD","NonDiscBal","NormalBal",column_name_operational,column_name_excess]]
#result_tex = result.applymap(billions)
##result.columns = ['DAILY_SPOT_BAL_USD','OPERATIONAL_BAL_USD','EXCESS_BAL_USD']
#result_tex = result_tex.sort_index(ascending=False)
#result_tex.to_latex(plot_folder+"/Summary_Plot2.tex",longtable=True)
#
#data_all_hedge_fund=[]
#data_all_miss_rai=[]
#data_all_YND=[]
#data_all_NND=[]
#result_ynd=[]
#result_nnd=[]
#result=[]
#result_hedge_fund=[]
#gc.collect()

##Decompose for Summary
# describe hedge fund excluded
def aggregate_summary_table_format_combined(data_all,currentsnapshot,benchmarkdate,attribute_groupby,attribute_agg,need_aggregate=False,group_agg=None,group_name=None,top_number=7):
#    data_all = data_all_tmp
#    attribute_groupby = attribute_groupby
#    attribute_agg = column_name_spot
#    group_agg = Alternatives
    if benchmarkdate.__class__ != list:
        benchmarkdate_tmp = [benchmarkdate]
    else:
        benchmarkdate_tmp = benchmarkdate
        
    if currentsnapshot.__class__ != list:
        currentsnapshot_tmp = [currentsnapshot]
    else:
        currentsnapshot_tmp = currentsnapshot
    
    #get average of as of period and benchmark period
    data_date = data_all[data_all['AS_OF_DATE'].isin(currentsnapshot_tmp + benchmarkdate_tmp)]
    table_orig=summarizeTable(data_date,attribute_groupby,attribute_agg).sort_values(currentsnapshot_tmp,ascending=False)
    table_all = table_orig[currentsnapshot_tmp].apply(np.nanmean,axis=1).to_frame("AsOfPeriod")
    table_all["Benchmark"] = table_orig[benchmarkdate_tmp].apply(np.nanmean,axis=1)

    if need_aggregate:
        if group_agg == None:
            is_aggregate = False
            if len(table_all)>top_number:
                table_all["DiffToBck"]=(table_all.iloc[:,0]-table_all.iloc[:,1])
                top_category = table_all.sort_values("DiffToBck",ascending=False).head(top_number).index
#                top_category = table_all.apply(np.nansum,axis=1).sort_values(ascending=False).head(top_number).index
                del table_all["DiffToBck"]
                is_aggregate = True
            if is_aggregate:
                other = table_all[~table_all.index.isin(top_category)].sum()
                other = other.to_frame(name="Other").transpose()
                table_all = table_all[table_all.index.isin(top_category)].append(other)
                table_all.sort_values(table_all.columns[-1],ascending=False,inplace=True)
        else:
            subtotal = table_all[table_all.index.isin(group_agg)].apply(np.nansum,axis=0)
            if group_name == None:
                subtotal.name = "Other"
            else:
                subtotal.name = group_name
            table_all = table_all[~table_all.index.isin(group_agg)].append(subtotal).sort_values("AsOfPeriod",ascending=False)
                
    table_all["DiffToBck"]=(table_all.iloc[:,0]-table_all.iloc[:,1])
    table_all = table_all.sort_values(["DiffToBck"],ascending=False)
    indexnames = table_all.index.tolist()
    del table_all["DiffToBck"]
    table_all = table_all.append(table_all.apply(np.nansum,axis=0),ignore_index=True)
    table_all.index = indexnames+["Total"]
    table_all["DiffToBck"]=(table_all.iloc[:,0]-table_all.iloc[:,1])
    return(table_all[["AsOfPeriod","Benchmark","DiffToBck"]])

#####Decompose for summary by Product, LE, Currency, BU, Region
def get_closest_date(result,currentdate, numdaysago):
#    result = data_all
    result_d = result.columns - (currentdate-datetime.timedelta(days=numdaysago))
    if any(result_d.days<=0):
        result_d = abs(result_d[result_d.days<=0])
        
    result_d = result_d==min(result_d)
    column = [i for i, x in enumerate(result_d) if x]
    return(result.columns[column][0])

def aggregate_summary_table(data_all,currentsnapshot,benchmarkdate,attribute_groupby,attribute_agg,need_aggregate=False,group_agg=None,group_name=None,top_number=7,sort_by="DiffToBck"):
#    data_all = data_all_tmp
#    attribute_groupby = attribute_groupby
#    attribute_agg = column_name_spot
#    group_agg = Alternatives
    if benchmarkdate.__class__ != list:
        benchmarkdate_tmp = [benchmarkdate]
    else:
        benchmarkdate_tmp = benchmarkdate
        
    if currentsnapshot.__class__ != list:
        currentsnapshot_tmp = [currentsnapshot]
    else:
        currentsnapshot_tmp = currentsnapshot
    
    #get average of as of period and benchmark period
    data_date = data_all[data_all['AS_OF_DATE'].isin(currentsnapshot_tmp + benchmarkdate_tmp)]
    table_orig=summarizeTable(data_date,attribute_groupby,attribute_agg).sort_values(currentsnapshot_tmp,ascending=False)
    table_all = table_orig[currentsnapshot_tmp].apply(np.nanmean,axis=1).to_frame("AsOfPeriod")
    table_all["Benchmark"] = table_orig[benchmarkdate_tmp].apply(np.nanmean,axis=1)
    
    #get average of one month time window. The rolling windowns starts at the maximum of as of period and look back 30 days.
    result = summarizeTable(data_all,attribute_groupby,attribute_agg)
    onemonthago = get_closest_date(result,max(currentsnapshot_tmp), 30)
    month_dates = [max(currentsnapshot_tmp) - datetime.timedelta(days=x) for x in range((max(currentsnapshot_tmp)-onemonthago).days+1)]
    result_one_month = result[result.columns[result.columns.isin(month_dates)]]
    average = result_one_month.apply(np.nanmean,axis=1)
    table_all["MonthAvg"] = average

    if need_aggregate:
        if group_agg == None:
            is_aggregate = False
            if len(table_all)>top_number:
                table_all["DiffToBck"]=(table_all.iloc[:,0]-table_all.iloc[:,1])
                top_category = table_all.sort_values(sort_by,ascending=False).head(top_number).index
#                top_category = table_all.apply(np.nansum,axis=1).sort_values(ascending=False).head(top_number).index
                del table_all["DiffToBck"]
                is_aggregate = True
            if is_aggregate:
                other = table_all[~table_all.index.isin(top_category)].sum()
                other = other.to_frame(name="Other").transpose()
                table_all = table_all[table_all.index.isin(top_category)].append(other)
#                table_all.sort_values(table_all.columns[-1],ascending=False,inplace=True)
        else:
            subtotal = table_all[table_all.index.isin(group_agg)].apply(np.nansum,axis=0)
            if group_name == None:
                subtotal.name = "Other"
            else:
                subtotal.name = group_name
            table_all = table_all[~table_all.index.isin(group_agg)].append(subtotal).sort_values("AsOfPeriod",ascending=False)
                
    table_all["DiffToBck"]=(table_all.iloc[:,0]-table_all.iloc[:,1])
    table_all = table_all.sort_values(sort_by,ascending=False)
 
    if "UnregulatedFund" in table_all.index:
        other = table_all.ix["UnregulatedFund"]
        table_all = table_all[~table_all.index.isin(["UnregulatedFund"])]
        other.name = "UnregulatedFunds"
        table_all = table_all.append(other)

    if "Other" in table_all.index:
        other = table_all.ix["Other"]
        table_all = table_all[~table_all.index.isin(["Other"])]
        table_all = table_all.append(other)
    
    indexnames = table_all.index.tolist()
    del table_all["DiffToBck"]
    table_all = table_all.append(table_all.apply(np.nansum,axis=0),ignore_index=True)
    table_all.index = indexnames+["Total"]
    table_all["DiffToBck"]=(table_all.iloc[:,0]-table_all.iloc[:,1])
    table_all["DiffToAvg"]=(table_all.iloc[:,0]-table_all.iloc[:,2])
    table_all["%DiffToBck"]=((table_all.iloc[:,0]-table_all.iloc[:,1])/table_all.iloc[:,1]).map(percentages)
    table_all["%DiffToAvg"]=((table_all.iloc[:,0]-table_all.iloc[:,2])/table_all.iloc[:,2]).map(percentages)

    table_all.iloc[:,0:5]=round(table_all.iloc[:,0:5]/1e9,2)
    return(table_all[["AsOfPeriod","Benchmark","DiffToBck","%DiffToBck","MonthAvg","DiffToAvg","%DiffToAvg"]])

#pivot tables
def pivot_tables(data_all_tmp, index, pivot_attributes, today=today,yesterday=yesterday,func = np.nansum):
#    pivot_attributes = pivot_attributes_client
    tmp_data_all = data_all_tmp[data_all_tmp["AS_OF_DATE"].isin([today,yesterday])]
#    tmp_data_all.loc[[x not in ["SSBT USA","SSBT GmbH","SSBT London"] for x in tmp_data_all[column_name_company]],column_name_company] = "Others"
    tmp_data_all.loc[[x not in ["DDA-DOMESTIC","IBDDA-DOMESTIC","IBDDA-EUROPEAN",'IBDDA-CAYMAN','IBDDA-LONDON'] for x in tmp_data_all['ProductTypes']],'ProductTypes'] = "Others"
    
    if index == column_name_parent:
        tmp_data_all.loc[[x not in ["SSBT USA","SSBT GmbH","SSBT London"] for x in tmp_data_all[column_name_company]],column_name_company] = "Others"
#        tmp_data_all.loc[[x not in ["USD","EUR","GBP"] for x in tmp_data_all[column_name_currency]],column_name_currency] = "Others"

    for j in range(len(pivot_attributes)):
        data_all_today = tmp_data_all[tmp_data_all["AS_OF_DATE"]==today]
        data_all_yesterday = tmp_data_all[tmp_data_all["AS_OF_DATE"]==yesterday]
        summarypivot = pd.pivot_table(data_all_today, values=column_name_spot, index=attribute_groupby, columns=pivot_attributes[j], aggfunc=func, fill_value=None, margins=False, dropna=True, margins_name='All')-pd.pivot_table(data_all_yesterday, values=column_name_spot, index=attribute_groupby, columns=pivot_attributes[j], aggfunc=func, fill_value=None, margins=False, dropna=True, margins_name='All')
        if 'Others' in summarypivot.columns:
            summarypivot.insert(len(summarypivot.columns)-1,'Others',summarypivot.pop('Others'))
        summarypivot.ix['Sum']=summarypivot.apply(np.nansum,axis=0)
        summarypivot['Sum']=summarypivot.apply(np.nansum,axis=1)
#        summarypivot = summarypivot.fillna("")
        round(summarypivot/1e9,1).fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attributes[j]+".tex",column_format='l'+'r'*len(summarypivot.columns),longtable=True)
        round(summarypivot/1e9,1).fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attributes[j]+".xlsx")
        if j == 0:
            pivots = summarypivot
        else:
            pivots = pd.concat([pivots,summarypivot],axis=1)
    return(pivots)
###change by attributes
#data prep
data_all_tmp = data_all[attribute_aggs +["AS_OF_DATE","ND_IND",column_name_parent,column_name_group,column_name_company,column_name_currency,column_name_region,'ProductTypes']]
data_all_tmp.loc[data_all_tmp[column_name_parent]=="UnregulatedFund",column_name_group]="Excluded Funds"
data_all_tmp.loc[(pd.isnull(data_all_tmp[column_name_group])) & (data_all_tmp[column_name_parent]!="UnregulatedFund"),column_name_group]="Funds Miss Model Info"
data_all_tmp.loc[(~pd.isnull(data_all_tmp[column_name_group])) & (data_all_tmp[column_name_parent]!="UnregulatedFund") & (data_all_tmp["ND_IND"]=="Y"),column_name_group]="Non-Discretionary Funds"
data_all_tmp.loc[["GmbH" in x for x in data_all_tmp[column_name_company]],column_name_company] = 'SSBT GmbH'
data_all_tmp.loc[[x not in ["USD","AUD","CAD","CHF","EUR","GBP","JPY"] for x in data_all_tmp[column_name_currency]],column_name_currency] = "Others"
IBDDADomestic = ['IBDDA-MONEY MARKET','IBDDA-NOW ACCOUNTS','IBDDA-EXTERNAL']
data_all_tmp.loc[[x in IBDDADomestic for x in data_all_tmp['ProductTypes']],'ProductTypes'] = "IBDDA-DOMESTIC"
data_all_tmp.loc[[x not in ["DDA-DOMESTIC","IBDDA-CAYMAN","IBDDA-DOMESTIC","IBDDA-EUROPEAN","IBDDA-LONDON"] for x in data_all_tmp['ProductTypes']],'ProductTypes'] = "Others"

pivot_attributes = [column_name_group,column_name_company,column_name_currency,column_name_region,'ProductTypes']

##behavior groups
attribute_groupby = column_name_group
result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result = result.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","INTRA-DAY","Multi-Day Low","Multi-Day High","Total"])
result.index = ['Excluded Funds','Funds Miss Model Info','Non-Discretionary Funds','Intraday Clients','Multi-Day Low Clients','Multi-Day High Clients', "Total"]
result.columns.name = "Fund/Client Type"

result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
result_total['Total'] = result_total.apply(np.nansum,axis=1)
summarypage1 = aggregate_summary_table_with_percentiles(result_total)
summarypage1 = summarypage1.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","INTRA-DAY","Multi-Day Low","Multi-Day High", "Total"])
summarypage1.index = ['Excluded Funds','Funds Miss Model Info','Non-Discretionary Funds','Intraday Clients','Multi-Day Low Clients','Multi-Day High Clients', "Total"]
summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
summarypage1diff = summarypage1diff.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","INTRA-DAY","Multi-Day Low","Multi-Day High", "Total"])
summarypage1diff.index = ['Excluded Funds','Funds Miss Model Info','Non-Discretionary Funds','Intraday Clients','Multi-Day Low Clients','Multi-Day High Clients', "Total"]
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")

result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
summarypage1 = summarypage1.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","INTRA-DAY","Multi-Day Low","Multi-Day High", "Total"])
summarypage1.index = ['Excluded Funds','Funds Miss Model Info','Non-Discretionary Funds','Intraday Clients','Multi-Day Low Clients','Multi-Day High Clients', "Total"]
summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
summarypage1diff = summarypage1diff.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","INTRA-DAY","Multi-Day Low","Multi-Day High", "Total"])
summarypage1diff.index = ['Excluded Funds','Funds Miss Model Info','Non-Discretionary Funds','Intraday Clients','Multi-Day Low Clients','Multi-Day High Clients', "Total"]
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")

summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
summarypage1.iloc[:,:7] = summarypage1.iloc[:,:7].applymap(percentages)
summarypage1.to_latex(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".xlsx")

#weekly dashboard
result_total_today = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total_today.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_total_lastweek = aggregate_summary_table_format_combined(data_all_tmp,today,lastweek,attribute_groupby,column_name_spot)
result_total_lastweek.columns = [today.strftime("%m-%d")+"[Total]",lastweek.strftime("%m-%d")+"[Total]","change"]
result_total = pd.concat([result_total_today.iloc[:,0:2],result_total_lastweek.iloc[:,1:2]],axis=1)
                        
result_operational_today = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational_today.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result_operational_lastweek = aggregate_summary_table_format_combined(data_all_tmp,today,lastweek,attribute_groupby,column_name_operational)
result_operational_lastweek.columns = [today.strftime("%m-%d")+"[Opertnl]",lastweek.strftime("%m-%d")+"[Opertnl]","change"]
result_operational = pd.concat([result_operational_today.iloc[:,0:2],result_operational_lastweek.iloc[:,1:2]],axis=1)
                               
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result = result.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","INTRA-DAY","Multi-Day Low","Multi-Day High","Total"])
result.index = ['Excluded Funds','Funds Miss Model Info','Non-Discretionary Funds','Intraday Clients','Multi-Day Low Clients','Multi-Day High Clients', "Total"]
result.columns.name = "Fund/Client Type"

result.to_latex(plot_folder+"/Summary_Weekly_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Weekly_"+attribute_groupby+".xlsx")

#benchmark dates
result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total.rename({"INTRA-DAY":"Intraday Clients","Multi-Day Low":"Multi-Day Low Clients","Multi-Day High":"Multi-Day High Clients"},inplace=True)
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational.rename({"INTRA-DAY":"Intraday Clients","Multi-Day Low":"Multi-Day Low Clients","Multi-Day High":"Multi-Day High Clients"},inplace=True)
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

pivot_tables(data_all_tmp, index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))))

### Region
attribute_groupby = column_name_region

result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result = result.reindex(["AMERICAS","EMEA","APAC","UNKNOWN","_Missing","Total"])
result.columns.name = "Region"

result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
result_total['Total'] = result_total.apply(np.nansum,axis=1)
summarypage1 = aggregate_summary_table_with_percentiles(result_total)
summarypage1 = summarypage1.reindex(["AMERICAS","EMEA","APAC","UNKNOWN","_Missing","Total"])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
summarypage1diff = summarypage1diff.reindex(["AMERICAS","EMEA","APAC","UNKNOWN","_Missing","Total"])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")

result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")

summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
summarypage1.iloc[:,:7] = summarypage1.iloc[:,:7].applymap(percentages)
summarypage1.to_latex(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".xlsx")

#benchmark dates
result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))))

### Legal Entity
attribute_groupby = column_name_company

result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result = result.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
result.columns.name = attribute_groupby

result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
result_total['Total'] = result_total.apply(np.nansum,axis=1)
result_total = result_total[['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total']]
summarypage1 = aggregate_summary_table_with_percentiles(result_total)
summarypage1 = summarypage1.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
summarypage1diff = summarypage1diff.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")

result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
result_ope = result_ope[['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total']]
summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
summarypage1 = summarypage1.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
summarypage1diff = summarypage1diff.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")

summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
summarypage1.iloc[:,:7] = summarypage1.iloc[:,:7].applymap(percentages)
summarypage1.to_latex(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".xlsx")

#weekly dashboard
result_total_today = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total_today.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_total_lastweek = aggregate_summary_table_format_combined(data_all_tmp,today,lastweek,attribute_groupby,column_name_spot)
result_total_lastweek.columns = [today.strftime("%m-%d")+"[Total]",lastweek.strftime("%m-%d")+"[Total]","change"]
result_total = pd.concat([result_total_today.iloc[:,0:2],result_total_lastweek.iloc[:,1:2]],axis=1)
                        
result_operational_today = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational_today.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result_operational_lastweek = aggregate_summary_table_format_combined(data_all_tmp,today,lastweek,attribute_groupby,column_name_operational)
result_operational_lastweek.columns = [today.strftime("%m-%d")+"[Opertnl]",lastweek.strftime("%m-%d")+"[Opertnl]","change"]
result_operational = pd.concat([result_operational_today.iloc[:,0:2],result_operational_lastweek.iloc[:,1:2]],axis=1)
                               
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result = result.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
result.columns.name = attribute_groupby

result.to_latex(plot_folder+"/Summary_Weekly_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Weekly_"+attribute_groupby+".xlsx")

#benchmark dates
result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))))

### Currency
attribute_groupby = column_name_currency
result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result = result.reindex(['USD','EUR','GBP','JPY','AUD','CAD','CHF','Others','Total'])
result.columns.name = attribute_groupby
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
result_total['Total'] = result_total.apply(np.nansum,axis=1)
summarypage1 = aggregate_summary_table_with_percentiles(result_total)
summarypage1 = summarypage1.reindex(['USD','EUR','GBP','JPY','AUD','CAD','CHF','Others','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
summarypage1diff = summarypage1diff.reindex(['USD','EUR','GBP','JPY','AUD','CAD','CHF','Others','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")

result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
summarypage1 = summarypage1.reindex(['USD','EUR','GBP','JPY','AUD','CAD','CHF','Others','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
summarypage1diff = summarypage1diff.reindex(['USD','EUR','GBP','JPY','AUD','CAD','CHF','Others','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")

summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
summarypage1.iloc[:,:7] = summarypage1.iloc[:,:7].applymap(percentages)
summarypage1.to_latex(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".xlsx")

result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))))

### Product Types
attribute_groupby = 'ProductTypes'
data_all_tmp_producttype = data_all_tmp[data_all_tmp[attribute_groupby].isin(['DDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-DOMESTIC'])]
result_total = aggregate_summary_table_format_combined(data_all_tmp_producttype,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp_producttype,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result = result.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','Total'])
result.index = ['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','Total - SSBT USA']
result.columns.name = attribute_groupby
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
result_total['Total'] = result_total.apply(np.nansum,axis=1)
summarypage1 = aggregate_summary_table_with_percentiles(result_total)
summarypage1 = summarypage1.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
summarypage1diff = summarypage1diff.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")

result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
summarypage1 = summarypage1.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
summarypage1diff = summarypage1diff.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")

summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
summarypage1.iloc[:,:7] = summarypage1.iloc[:,:7].applymap(percentages)
summarypage1.to_latex(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".xlsx")

result_total = aggregate_summary_table(data_all_tmp_producttype,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total.index = [x.replace("Total","Total - SSBT USA") for x in result_total.index]
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp_producttype,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational.index = [x.replace("Total","Total - SSBT USA") for x in result_operational.index]
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

pivot_tables(data_all_tmp_producttype,index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))))
data_all_tmp_producttype = []

result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
#result = result.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','Total'])
#result.index = ['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','Total - SSBT USA']
result.columns.name = attribute_groupby
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_all.tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_all.xlsx")

result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total.index = [x.replace("Total","Total - SSBT USA") for x in result_total.index]
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+"_all.tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational.index = [x.replace("Total","Total - SSBT USA") for x in result_operational.index]
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+"_all.tex",column_format="l|r|rrr|rrr",longtable=True)

# subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis_'+day, './document.tex'])

## Parent Code ---Unique one
#increasing
attribute_groupby = column_name_parent
result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,2)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result = result[~result.index.isin(["UnregulatedFund","Total"])]
result = result.head(num_top_funds_to_plot)
result.index = [x.replace("_Missing","Fund Miss Model Info") for x in result.index]
client_largest_increasing = result.index
result = result.join(mapping_code_names,how="left")
null_client = pd.isnull(result["Client"])
result.loc[null_client,"Client"] = result[null_client].index
result.insert(0, "Client",result.pop("Client"))
result.index = list(range(1,len(result)+1))
result.to_latex(plot_folder+"/Summary_Intraday_Deposit_"+attribute_groupby+".tex",column_format="ll|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_Deposit_"+attribute_groupby+".xlsx")
#decreasing
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,2)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result = result.sort_values("change")
result = result[~result.index.isin(["UnregulatedFund","Total"])]
result = result.head(num_top_funds_to_plot)
result.index = [x.replace("_Missing","Fund Miss Model Info") for x in result.index]
client_largest_droping = result.index
result = result.join(mapping_code_names,how="left")
null_client = pd.isnull(result["Client"])
result.loc[null_client,"Client"] = result[null_client].index
result.insert(0, "Client",result.pop("Client"))
result.index = list(range(1,len(result)+1))
result.to_latex(plot_folder+"/Summary_Intraday_Withdraw_"+attribute_groupby+".tex",column_format="ll|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_Withdraw_"+attribute_groupby+".xlsx")

##Sort by difftobck
need_aggregate = True
top_number = num_top_funds_to_plot
result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot,need_aggregate=need_aggregate,top_number=top_number)
result_total.index = [x.replace("_Missing","Fund Miss Model Info") for x in result_total.index]
client_largest_average_droping = result_total.index
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="ll|r|rrr|rrr",longtable=True)

result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational,need_aggregate=need_aggregate,top_number=top_number)
result_operational.index = [x.replace("_Missing","Fund Miss Model Info") for x in result_operational.index]
result_operational = result_operational.join(mapping_code_names,how="left")
null_client = pd.isnull(result_operational["Client"])
result_operational.loc[null_client,"Client"] = result_operational[null_client].index
result_operational.insert(0, "Client",result_operational.pop("Client"))
result_operational.index = list(range(1,len(result_operational)-1))+["",""]
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="ll|r|rrr|rrr",longtable=True)

##sort by Average Balance
need_aggregate = True
top_number = num_top_funds_to_plot
sort_by = "MonthAvg"
result_total = aggregate_summary_table(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot,need_aggregate=need_aggregate,top_number=top_number,sort_by=sort_by)
top_clients = result_total.index[~result_total.index.isin(['_Missing','UnregulatedFunds', 'Others', 'Total','Other'])]
result_total.index = [x.replace("_Missing","Fund Miss Model Info") for x in result_total.index]
client_largest_average = result_total.index
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.columns = ['Client', today.strftime("%m-%d[Total]"), yesterday.strftime("%m-%d[Total]"), 'DiffToBck', '%DiffToBck','MonthAvg', 'DiffToAvg', '%DiffToAvg']
result_total.to_latex(plot_folder+"/Summary_Monthly_"+attribute_groupby+".tex",column_format="ll|r|rrr|rrr",longtable=True)

result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational,need_aggregate=need_aggregate,top_number=top_number)
result_operational.index = [x.replace("_Missing","Fund Miss Model Info") for x in result_operational.index]
result_operational = result_operational.join(mapping_code_names,how="left")
null_client = pd.isnull(result_operational["Client"])
result_operational.loc[null_client,"Client"] = result_operational[null_client].index
result_operational.insert(0, "Client",result_operational.pop("Client"))
result_operational.index = list(range(1,len(result_operational)-1))+["",""]
result_operational.columns = ['Client', today.strftime("%m-%d[Total]"), yesterday.strftime("%m-%d[Total]"), 'DiffToBck', '%DiffToBck','MonthAvg', 'DiffToAvg', '%DiffToAvg']
result_operational.to_latex(plot_folder+"/Summary_Monthly_Operational_"+attribute_groupby+".tex",column_format="ll|r|rrr|rrr",longtable=True)


## pivot table
#pivot_attributes = [column_name_group,column_name_company,column_name_currency,column_name_region,'ProductTypes']
data_all_tmp.loc[~data_all_tmp[column_name_parent].isin(top_clients),column_name_parent] = 'Others'

pivot_attribute = [column_name_company]
result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute)
result_total = round(result_total/1e9,2)
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
result_total = result_total.ix[list(top_clients)+["Others","Sum"]]
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".xlsx")
result_total.fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".tex",column_format='l'+'r'*len(result_total.columns),longtable=True)

pivot_attribute = [column_name_group]
result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute)
result_total = round(result_total/1e9,2)
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
result_total = result_total.ix[list(top_clients)+["Others","Sum"]]
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.columns = ['Client', 'Excluded', 'MissModelInfo', 'INTRADAY','MDHigh', 'MDLow', 'Non-Disc', 'Sum']
result_total.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".xlsx")
result_total.fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".tex",column_format='l'+'r'*len(result_total.columns),longtable=True)

pivot_attribute = [column_name_currency]
result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute)
result_total = round(result_total/1e9,2)
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
result_total = result_total.ix[list(top_clients)+["Others","Sum"]]
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".xlsx")
result_total.fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".tex",column_format='l'+'r'*len(result_total.columns),longtable=True)

pivot_attribute = [column_name_region]
result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute)
result_total = round(result_total/1e9,2)
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
result_total = result_total.ix[list(top_clients)+["Others","Sum"]]
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".xlsx")
result_total.fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".tex",column_format='l'+'r'*len(result_total.columns),longtable=True)

pivot_attribute = ['ProductTypes']
result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute)
result_total = round(result_total/1e9,2)
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
result_total = result_total.ix[list(top_clients)+["Others","Sum"]]
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.columns = ['Client', 'DDA/USA', 'IBDDA/CAYMAN', 'IBDDA/USA','IBDDA/EU', 'IBDDA/GBP', 'Others', 'Sum']
result_total.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".xlsx")
result_total.fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".tex",column_format='l'+'r'*len(result_total.columns),longtable=True)

data_all_tmp = []

###Country
domicile_diction = {"US":"USA",
                    "UNK":"GREAT BRITAIN",
                    "LU":"LUXEMBOURG",
                    "IE":"IRELAND",
                    "DK":"DENMARK",
                    "GB":"GREAT BRITAIN",
                    "DE":"GERMANY",
                    "AU":"AUSTRALIA",
                    "CA":"CANADA",
                    "CN":"CHINA",
                    "KY":"CAYMAN ISLANDS",
                    "CH":"SWITZERLAND",
                    "TW":"TAIWAN",
                    "JP":"JAPAN",
                    "SG":"SINGAPORE",
                    "BB":"BARBADOS",
                    "NL":"NETHERLAND",
                    "AT":"AUSTRIA",
                    "KW":"KUWAIT",
                    "KR":"KOREA",
                    "BN":"BRUNEI DARUSSALAM",
                    "FI":"FINLAND",
                    "HK":"HONG KONG",
                    "AE":"UAE",
                    "SA":"SAUDI ARABIA",
                    "BM":"BERMUDA",
                    "FR":"FRANCE",
                    "CL":"CHILE",
                    "Other":"Other",
                    "Total":"Total"}

attribute_groupby = 'COUNTRY_DOMICILE'
data_all_tmp = data_all[attribute_aggs+[attribute_groupby,"AS_OF_DATE"]]
data_all_tmp.loc[[x not in ['US','UNK','LU','IE','DK','GB','DE','AU','CA','CN','KY','JP'] for x in data_all_tmp[attribute_groupby]],attribute_groupby] = "Other"

result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result.index = [domicile_diction[x] for x in result.index]
result.columns.name = "Domicile"
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total.index = [domicile_diction[x] for x in result_total.index]
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational.index = [domicile_diction[x] for x in result_operational.index]
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

data_all_tmp = []

###Investment Style
attribute_groupby = column_name_stylegroup
data_all_tmp = data_all[attribute_aggs+[attribute_groupby,column_name_style,"AS_OF_DATE"]]
data_all_tmp.loc[["MONEY" in x for x in data_all_tmp[column_name_style].fillna("_Missing")],attribute_groupby] = "CASH - MONEY MARKETS"
Alternatives = ['CORPORATE - GENERAL','UNKNOWN_STT','INCOME MIXED','REAL ESTATE INVESTMENTS','OTHER_PWC','DERIVATIVES','PRIVATE PLACEMENTS']
data_all_tmp.loc[[x in Alternatives for x in data_all_tmp[column_name_stylegroup].fillna("_Missing")],column_name_stylegroup] = "ALTERNATIVE"

result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result.columns.name = attribute_groupby
result = result.reindex(["CASH","CASH - MONEY MARKETS","DOMESTIC BOND","GLOBAL BOND","BALANCED","DOMESTIC EQUITY","GLOBAL EQUITY","GROWTH","ALTERNATIVE","MUTUAL FUNDS HOLDING","_Missing","Total"])
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

data_all_tmp = []

###Business units
attribute_groupby = column_name_division
data_all_tmp = data_all[attribute_aggs+[attribute_groupby,"AS_OF_DATE"]]

result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result.iloc[:,3]/result.iloc[:,0]).map(percentages)
result.columns.name = attribute_groupby
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

data_all_tmp = []

#########################################################
aggregate_groupby = column_name_parent
result = summarizeTable(data_all,aggregate_groupby,column_name_spot)
result_operational = summarizeTable(data_all,aggregate_groupby,column_name_operational)

result_byid = summarizeTable(data_all,"ID",column_name_spot)
result_operational_byid = summarizeTable(data_all,"ID",column_name_operational)

#get attributes
mapping_number_behavior = data_all[["AS_OF_DATE","ID",column_name_parent,column_name_parentname,column_name_region,column_name_group,'ND_IND']].copy()
mapping_number_behavior[column_name_parentname] = [x.replace("(PARENT)","").replace("CONFIDENTIAL CLIENT","Conf. Client")[:34] for x in mapping_number_behavior[column_name_parentname].fillna("_Missing")]

table_name = "TotalDepositTableByParent"
total_deposit_by_parent = read_db(table_name)
table_name = "TotalDepositTableByParentRegion"
total_deposit_by_parent_region = read_db(table_name)
table_name = "OperationalDepositTableByParent"
operational_deposit_by_parent = read_db(table_name)
table_name = "OperationalDepositTableByParentRegion"
operational_deposit_by_parent_region = read_db(table_name)

table_name = "BindingConstraintTableByParent"
binding_constraint_by_parent = read_db(table_name)
table_name = "BindingConstraintTableByParentRegion"
binding_constraint_by_parent_region = read_db(table_name)

table_name = "BehaviorGroupTableByParentRegion"
behavior_group_by_parent_region = read_db(table_name)

table_name = "AveragePaymentTableByParentRegion"
average_payment_by_parent_region = read_db(table_name)

behavior_color_map = {"INTRA-DAY":"lightcoral","Multi-Day High":"deepskyblue","Multi-Day Low":"yellow"}

#mapping_number_behavior = data_all[["AS_OF_DATE","ID","CDMS_ULT_PRNT_CODE",column_name_region,\
#                                    "ULT_PARENT_NAME"]].copy()

#mapping_number_behavior.rename({"CDMS_ULT_PRNT_CODE":"ParentCo","BEHAVIORAL_GROUP":"Group"})

#filter those with largest average balance
#numdaysago = 30
#currentdate = currentsnapshot
#def get_closest_date(result,currentdate, numdaysago):
#    result_d = result.columns - (currentdate-datetime.timedelta(days=numdaysago))
#    if any(result_d.days<=0):
#        result_d = abs(result_d[result_d.days<=0])
#        
#    result_d = result_d==min(result_d)
#    column = [i for i, x in enumerate(result_d) if x]
#    return(result.columns[column][0])

#yesterday = get_closest_date(result,currentsnapshot, 1)
#onemonthago = get_closest_date(result,currentsnapshot, 30)
#summary_clients = result[[currentsnapshot,yesterday]]
#summary_clients["DailyChg"] = summary_clients[currentsnapshot]-summary_clients[yesterday]
#month_dates = [currentsnapshot - datetime.timedelta(days=x) for x in range((currentsnapshot-onemonthago).days+1)]
#summary_clients['MonthAvg'] = result[result.columns[result.columns.isin(month_dates)]].apply(np.nanmean,axis=1)
#summary_clients['DiffToAvg'] = summary_clients[currentsnapshot] - summary_clients['MonthAvg']
#
#########Largest average balance
#summary_the_largest = summary_clients[~summary_clients.index.isin(["UnregulatedFund","_Missing"])].sort_values(by=['MonthAvg'],ascending=False).head(num_top_funds_to_plot*3)
##client_largest_average = summary_the_largest.index
#summary_others = summary_clients[~summary_clients.index.isin(summary_the_largest.index)]
#subtotal_included = summary_the_largest.apply(np.nansum,axis=0)
#subtotal_included.name = "Subtotal-Top Clients"
#subtotal_excluded = summary_others.apply(np.nansum,axis=0)
#subtotal_excluded.name = "Subtotal-Others"
#subtotal_all = summary_clients.apply(np.nansum,axis=0)
#subtotal_all.name = "Total"
#summary_the_largest = summary_the_largest.append(subtotal_included)
#summary_the_largest = summary_the_largest.append(subtotal_excluded)
#summary_the_largest = summary_the_largest.append(subtotal_all)
#summary_the_largest = summary_the_largest.join(mapping_code_names,how="left")
#summary_the_largest.insert(0, "Client",summary_the_largest.pop(column_name_parentname))
#summary_the_largest.insert(0, "Code", summary_the_largest.index)
#summary_the_largest.index = list(range(1,len(summary_the_largest)-2))+["","",""]
#
#summary_the_largest.columns = [["Code","Client",currentsnapshot.strftime("%Y-%m-%d"),yesterday.strftime("%Y-%m-%d"),"DailyChg","MonthAvg","DiffToAvg"]]
#summary_the_largest.to_excel(plot_folder+"/"+aggregate_groupby+"_LargestAveBalance.xlsx")
#
#summary_the_largest['Client'][pd.isnull(summary_the_largest['Client'])] = ""
#summary_the_largest[[currentsnapshot.strftime("%Y-%m-%d"),yesterday.strftime("%Y-%m-%d"),"DailyChg","MonthAvg","DiffToAvg"]] = summary_the_largest[[currentsnapshot.strftime("%Y-%m-%d"),yesterday.strftime("%Y-%m-%d"),"DailyChg","MonthAvg","DiffToAvg"]].applymap(millions)
#summary_the_largest.to_latex(plot_folder+"/"+aggregate_groupby+"_LargestAveBalance.tex",column_format="llrrrrrr",longtable=True)
#
#
#
#########Largest dropping
#summary_the_largest = summary_clients[~summary_clients.index.isin(["UnregulatedFund","_Missing"])].sort_values(by=['DiffToAvg'],ascending=True).head(num_top_funds_to_plot*3)
##client_largest_droping = summary_the_largest.index
#summary_others = summary_clients[~summary_clients.index.isin(summary_the_largest.index)]
#subtotal_included = summary_the_largest.apply(np.nansum,axis=0)
#subtotal_included.name = "Subtotal-Top Clients"
#subtotal_excluded = summary_others.apply(np.nansum,axis=0)
#subtotal_excluded.name = "Subtotal-Others"
#subtotal_all = summary_clients.apply(np.nansum,axis=0)
#subtotal_all.name = "Total"
#summary_the_largest = summary_the_largest.append(subtotal_included)
#summary_the_largest = summary_the_largest.append(subtotal_excluded)
#summary_the_largest = summary_the_largest.append(subtotal_all)
#summary_the_largest = summary_the_largest.join(mapping_code_names,how="left")
#summary_the_largest.insert(0, "Client",summary_the_largest.pop(column_name_parentname))
#summary_the_largest.insert(0,"Code",summary_the_largest.index)
#summary_the_largest.index = list(range(1,len(summary_the_largest)-2))+["","",""]
#
#summary_the_largest.columns = [["Code","Client",currentsnapshot.strftime("%Y-%m-%d"),yesterday.strftime("%Y-%m-%d"),"DailyChg","MonthAvg","DiffToAvg"]]
#summary_the_largest.to_excel(plot_folder+"/"+aggregate_groupby+"_LargestWithdraw_Summary.xlsx")
#
#summary_the_largest['Client'][pd.isnull(summary_the_largest['Client'])] = ""
#summary_the_largest[[currentsnapshot.strftime("%Y-%m-%d"),yesterday.strftime("%Y-%m-%d"),"DailyChg","MonthAvg","DiffToAvg"]] = summary_the_largest[[currentsnapshot.strftime("%Y-%m-%d"),yesterday.strftime("%Y-%m-%d"),"DailyChg","MonthAvg","DiffToAvg"]].applymap(millions)
#summary_the_largest.to_latex(plot_folder+"/"+aggregate_groupby+"_LargestWithdraw_Summary.tex",column_format="llrrrrrr",longtable=True)
#
#########Largest increasing
#summary_the_largest = summary_clients[~summary_clients.index.isin(["UnregulatedFund","_Missing"])].sort_values(by=['DiffToAvg'],ascending=False).head(num_top_funds_to_plot*3)
##client_largest_increasing = summary_the_largest.index
#summary_others = summary_clients[~summary_clients.index.isin(summary_the_largest.index)]
#subtotal_included = summary_the_largest.apply(np.nansum,axis=0)
#subtotal_included.name = "Subtotal-Top Clients"
#subtotal_excluded = summary_others.apply(np.nansum,axis=0)
#subtotal_excluded.name = "Subtotal-Others"
#subtotal_all = summary_clients.apply(np.nansum,axis=0)
#subtotal_all.name = "Total"
#summary_the_largest = summary_the_largest.append(subtotal_included)
#summary_the_largest = summary_the_largest.append(subtotal_excluded)
#summary_the_largest = summary_the_largest.append(subtotal_all)
#summary_the_largest = summary_the_largest.join(mapping_code_names,how="left")
#summary_the_largest.insert(0, "Client",summary_the_largest.pop(column_name_parentname))
#summary_the_largest.insert(0,"Code",summary_the_largest.index)
#summary_the_largest.index = list(range(1,len(summary_the_largest)-2))+["","",""]
#
#summary_the_largest.columns = [["Code","Client",currentsnapshot.strftime("%Y-%m-%d"),yesterday.strftime("%Y-%m-%d"),"DailyChg","MonthAvg","DiffToAvg"]]
#summary_the_largest.to_excel(plot_folder+"/"+aggregate_groupby+"_LargestDeposit_Summary.xlsx")
#
#summary_the_largest['Client'][pd.isnull(summary_the_largest['Client'])] = ""
#summary_the_largest[[currentsnapshot.strftime("%Y-%m-%d"),yesterday.strftime("%Y-%m-%d"),"DailyChg","MonthAvg","DiffToAvg"]] = summary_the_largest[[currentsnapshot.strftime("%Y-%m-%d"),yesterday.strftime("%Y-%m-%d"),"DailyChg","MonthAvg","DiffToAvg"]].applymap(millions)
#summary_the_largest.to_latex(plot_folder+"/"+aggregate_groupby+"_LargestDeposit_Summary.tex",column_format="llrrrrrr",longtable=True)

###Trend of top funds selected
def transform_string_to_legend(string,code):
    return(string.replace(code,code+"/")[:-1]+"/"+string.replace(code,code+"/")[-1])
def extend_ts(spot,total_deposit_by_parent):
    spot_extended = total_deposit_by_parent.ix[spot.columns].transpose()
    flag_is_identical = all(abs(spot_extended.ix[spot.index]-spot.transpose()) < 1)
    if flag_is_identical:
        return(spot_extended[spot_extended.index <= max(spot.index)])
    else:
        return(spot)

def graph_fund_list(fund_list,plot_threshold=True):
    fund_list = fund_list[~pd.isnull(fund_list)]
    fund_list = fund_list[~fund_list.isin(['Fund Miss Model Info','UnregulatedFunds','Other'])]
    
    for i in range(len(fund_list)):
#        i = 1
        spot = result[result.index==fund_list[i]].transpose()
        spot = extend_ts(spot,total_deposit_by_parent)
        spot.columns = spot.columns+": Spot Balance"
        ope = result_operational[result_operational.index==fund_list[i]].transpose()
        ope = extend_ts(ope,operational_deposit_by_parent)
        ope.columns = ope.columns+": Operational Balance"
        statistics = spot.describe(percentiles=[0.05,0.95])
        stat_mean = statistics.loc["mean",:][0]
        stat_5per = statistics.loc["5%",:][0]
        stat_95per = statistics.loc["95%",:][0]
        stat_min = statistics.loc["min",:][0]
        stat_max = statistics.loc["max",:][0]

        constraint = binding_constraint_by_parent.ix[[fund_list[i]]].transpose()[binding_constraint_by_parent.columns <= max(spot.index)]
        
        fig, ax = plt.subplots()
        spot.dropna().plot(ax=ax,style=styles)
        ope.dropna().plot(ax=ax,style=styles2)
        constraint.plot(kind = "area",color="green",ax=ax,alpha=0.1,label=None,legend=False)
        ax.axhline(y=stat_mean, color='r', linestyle="--",linewidth=2)
        ax.axhline(y=stat_5per, color='r', linestyle="-.",linewidth=1.5)
        ax.axhline(y=stat_95per, color='r', linestyle="-.",linewidth=1.5)
        ax.axhline(y=stat_min, color='r', linestyle=":",linewidth=1)
        ax.axhline(y=stat_max, color='r', linestyle=":",linewidth=1)
        ax.yaxis.set_major_formatter(formatter_billions)
        ax.tick_params(labelsize=20)
        ax.xaxis.label.set_visible(False)
        ax.set_title("with Mean, Max, Min, 5th and 95th Percentiles of Total Deposits in red dotted lines",fontsize=15)
        
        labels = spot.columns.append(ope.columns)
        lines, _ = ax.get_legend_handles_labels()
        ax.legend(lines, labels, loc='best')
        fig.suptitle("Trend of Deposit Balance of " + fund_list[i] + " (in USD)", fontsize=20)
        fig.set_size_inches(12,9)
        
        fig.savefig(plot_folder+"/Trend_"+aggregate_groupby+"_"+fund_list[i]+".png",bbox_inches='tight')
        plt.clf()
        
        subs = mapping_number_behavior[mapping_number_behavior[column_name_parent]==fund_list[i]]
        subs = subs.dropna()
        subs_unique = subs.drop_duplicates(["ID",column_name_group ,column_name_parentname, column_name_parent])
        subs_unique = subs_unique.sort_values(["ID"])
        del subs_unique['ID']

        if len(subs_unique) > 0:
#            subs.to_latex(plot_folder+"/Trend_"+aggregate_groupby+"_"+fund_list[i]+"_details.tex",index=False,longtable=True)
            code = subs_unique[column_name_parent].drop_duplicates()[0]
            funds = subs_unique.index.unique()
            funds = funds[~pd.isnull(funds)]

            # analysis of entire clients' deposit base
            spot = result_byid.loc[funds].transpose()
            spot = extend_ts(spot,total_deposit_by_parent_region)
            subtotal = spot.apply(np.nansum,axis=1)
            subtotal.name = "Total"
            total = pd.concat([spot,subtotal],axis=1).sort_index(ascending=False)
            summarypage_spot = aggregate_summary_table_with_percentiles(total,attributes=total.columns,digit=2)
            summarypage_spot = summarypage_spot.join(subs_unique[['BehaviorGroup']],how="left")
            summarypage_spot.insert(0,"BehaviorGroup",summarypage_spot.pop("BehaviorGroup"))
            if fund_list[i] in mapping_code_names.index:
                client_name = mapping_code_names.ix[fund_list[i]][0]
            else:
                client_name = fund_list[i]
            
            summarypage_spot.insert(0,'Client',client_name)
            summarypage_spot.index = [transform_string_to_legend(x,code) if x[-1] in ["Y","N"] else x for x in summarypage_spot.index]
            summarypage_spot.to_latex(plot_folder+"/Trend_"+aggregate_groupby+"_Total_"+fund_list[i]+".tex",longtable=True)
            
            ope = result_operational_byid.loc[funds].transpose()
            ope = extend_ts(ope,operational_deposit_by_parent_region)
            subtotal = ope.apply(np.nansum,axis=1)
            subtotal.name = "Total"
            total = pd.concat([ope,subtotal],axis=1).sort_index(ascending=False)
            summarypage_ope = aggregate_summary_table_with_percentiles(total,attributes=total.columns,digit=2)
            summarypage_ope = summarypage_ope.join(subs_unique[['BehaviorGroup']],how="left")
            summarypage_ope.insert(0,"BehaviorGroup",summarypage_ope.pop("BehaviorGroup"))

            summarypage_ope.insert(0,'Client',client_name)
            summarypage_ope.index = [transform_string_to_legend(x,code) if x[-1] in ["Y","N"] else x for x in summarypage_ope.index]
            summarypage_ope.to_latex(plot_folder+"/Trend_"+aggregate_groupby+"_Operational_"+fund_list[i]+".tex",longtable=True)
            
            # get funds that represent 5% or more of total client
            share = spot.apply(np.nansum,axis=0)/np.nansum(spot)
            funds = share[share>0.05].index

            spot = result_byid.loc[funds].transpose()
            spot = extend_ts(spot,total_deposit_by_parent_region)
            ope = result_operational_byid.loc[funds].transpose()
            ope = extend_ts(ope,operational_deposit_by_parent_region)
            
            if len(ope.columns) > 0:
                
                constraints = data_all.loc[data_all['ID'].isin(funds),["AS_OF_DATE","OPERATIONAL_BINDING_CONSTRAINT","ID"]]
                constraints.index = range(len(constraints))        
                constraints.drop_duplicates(inplace=True)
                constraints = constraints.pivot(index='AS_OF_DATE', columns='ID', values='OPERATIONAL_BINDING_CONSTRAINT')
                constraints = extend_ts(constraints,binding_constraint_by_parent_region)
                constraints = constraints.stack().to_frame()
                constraints.reset_index(inplace=True)
                constraints.columns = ["AS_OF_DATE","ID","OPERATIONAL_BINDING_CONSTRAINT"]
                constraints.set_index('AS_OF_DATE',inplace=True)
                
                #                add behavior group changes using subs!!!!!!!
#                subs[["AS_OF_DATE","ID","BehaviorGroup"]].to_excel(plot_folder+"/test.xlsx")
#                subs.to_excel(plot_folder+"/test.xlsx")

                behaviors = subs[["AS_OF_DATE","ID","BehaviorGroup"]].drop_duplicates().pivot(index="AS_OF_DATE",columns="ID",values="BehaviorGroup")
                behaviors = extend_ts(behaviors,behavior_group_by_parent_region)
                
                fig, ax = plt.subplots()
                spot.plot(ax=ax, style=styles)
                ope.plot(ax=ax, style=styles2)
                constraints.groupby("ID").plot(kind = "area",color="green",ax=ax,alpha=0.1,label=None,legend=False)
#                behaviors
                bandwidth = np.nanmax(spot)/30
                for j in range(len(spot.columns)):
                    beha = behaviors[[spot.columns[j]]]
                    beha.rename(columns={spot.columns[j]:"BehaviorGroups"},inplace=True)
                    groups = spot.iloc[:,j:(j+1)].join(beha,how="left")
                    groups = groups.dropna()
                    last_row = 0
                    y = []
                    for row in range(1,len(groups)):
                        if groups.iloc[row,1]!=groups.iloc[row-1,1]:
                            y = groups.iloc[last_row:row,:][spot.columns[j]]
                            last_row = row
                        elif row == len(groups)-1:
                            y = groups.iloc[last_row:,:][spot.columns[j]]
                        if len(y) > 0:
                            ax.fill_between(y.index, y-bandwidth, y+bandwidth, facecolor=behavior_color_map[groups.iloc[row-1,1]], alpha=0.5)
                            y = []
                #legends
                labels = [(transform_string_to_legend(x,code)+":TotalBal") for x in spot.columns]+[(transform_string_to_legend(x,code)+ ":OpertnlBal") for x in ope.columns]
                lines, _ = ax.get_legend_handles_labels()
                ax.legend(lines, labels, loc='best')
                ax.yaxis.set_major_formatter(formatter_billions)
                ax.tick_params(labelsize=20)
                ax.xaxis.label.set_visible(False)
                plt.title("Trend of Spot and Operational Deposit of " + fund_list[i])
                fig.set_size_inches(12,9)
                
                fig.savefig(plot_folder+"/Trend_"+aggregate_groupby+"_"+fund_list[i]+"_details.png",bbox_inches='tight')
                plt.clf()

#fund_list = client_largest_average_droping[:int(num_top_funds_to_plot)]
#graph_fund_list(fund_list, plot_threshold = False)

fund_list = client_largest_average[:num_top_funds_to_plot]
graph_fund_list(fund_list, plot_threshold = False)

#fund_list = client_largest_droping[:int(num_top_funds_to_plot)]
#graph_fund_list(fund_list, plot_threshold = False)
#
#fund_list = client_largest_increasing[:int(num_top_funds_to_plot)]
#graph_fund_list(fund_list, plot_threshold = False)

##get fund list for largest funds
if benchmarkdate.__class__ != list:
    benchmarkdatetex = [benchmarkdate]
else:
    benchmarkdatetex = benchmarkdate
    
if currentsnapshot.__class__ != list:
    currentsnapshottex = [currentsnapshot]
else:
    currentsnapshottex = currentsnapshot

file = open(pdf_dir + "/GlobalParameters.tex", "w")
command="\\newcommand\\Output{"+plot_folder+"}\n"+\
"\\newcommand\\summaryLE{"+column_name_company +"}\n"+\
"\\newcommand\\summaryBehavior{"+column_name_group +"}\n"+\
"\\newcommand\\summaryCurrency{"+column_name_currency +"}\n"+\
"\\newcommand\\summaryDomicile{"+'COUNTRY_DOMICILE'+"}\n"+\
"\\newcommand\\summaryProduct{"+'ProductTypes'+"}\n"+\
"\\newcommand\\summaryClient{"+column_name_parent +"}\n"+\
"\\newcommand\\summaryStyle{"+column_name_stylegroup+"}\n"+\
"\\newcommand\\summaryBU{"+column_name_division +"}\n"+\
"\\newcommand\\summarySegment{"+column_name_market+"}\n"+\
"\\newcommand\\summaryRegion{"+column_name_region+"}\n"+\
"\\newcommand\\PivotAttributes{"+", ".join(pivot_attributes)+"}\n"+\
"\\newcommand\\DepositTypes{"+",".join(attribute_aggs)+"}\n"+\
"\\newcommand\\Operational{"+column_name_operational+"}\n"+\
"\\newcommand\\Excess{"+column_name_excess+"}\n"+\
"\\newcommand\\Total{"+column_name_spot+"}\n"+\
"\\newcommand\\BenchmarkDate{"+"from "+min(benchmarkdatetex).strftime("%Y-%m-%d") +" to " + max(benchmarkdatetex).strftime("%Y-%m-%d")+"}\n"+\
 "\\newcommand\\CurrentDate{"+", ".join([x.strftime("%Y-%m-%d") for x in currentsnapshottex])+"}\n"+\
"\\newcommand\\TopParentDailyWidthdraw{"+",".join(client_largest_droping[:num_top_funds_to_plot])+"}\n"+\
"\\newcommand\\TopParentDailyDeposit{"+",".join(client_largest_increasing[:num_top_funds_to_plot])+"}\n"+\
"\\newcommand\\TopParentMonthlyAverageBalance{"+",".join(client_largest_average[:num_top_funds_to_plot])+"}\n"+\
"\\newcommand\\TopParentMonthlyWithdraw{"+",".join(client_largest_average_droping[:num_top_funds_to_plot])+"}\n"+\
"\\newcommand\\TopParent{"+column_name_parent+"}\n"
file.write(command)
file.close()

os.chdir(pdf_dir)
#subprocess.call(['pdflatex', '-output-directory', "./", '-jobname', 'DepositAnalysis_'+day, './document.tex'])

subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis_'+day, './document.tex'])
subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis_'+day, './document.tex'])
if PRODUCTION_ENVRIONMENT:
    subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis_'+day, './document.tex'])
    subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis_'+day, './document.tex'])

subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_'+day, './document_workingcopy.tex'])
subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_'+day, './document_workingcopy.tex'])
if PRODUCTION_ENVRIONMENT:
    subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_'+day, './document_workingcopy.tex'])
    subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_'+day, './document_workingcopy.tex'])

end_time = time.time()
print("Execution runtime is "+str(round(end_time-start_time,0))+" seconds")

