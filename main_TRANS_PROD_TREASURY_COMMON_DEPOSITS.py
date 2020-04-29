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
    weekly_table_dir = "Z:/Liquidity Risk/Operational Deposits/WeeklyReportTable"
    num_top_funds_to_plot = 40
    TEST_ENVIRONMENT = False
    FED_ENVIRONMENT = False
else:
    code_dir = "Z:/Charles/ORM/SourceCodes"
    output_dir = "Z:/Charles/ORM/Output_Test"
    pdf_dir = "Z:/Charles/ORM/PDFReport_Test"
    weekly_table_dir = output_dir
    num_top_funds_to_plot = 40
    TEST_ENVIRONMENT = True
    FED_ENVIRONMENT = False

mycommontable_dir = "Z:/FTDRDataBase/OPS_DEPOSIT_PBUP_FBO_DDA_GTRM"
input_file = "Z:/Charles/ORM/Data/Parameters.xlsx"
daily_operational_balance_input_file = "Z:/Charles/ORM/Data/DAILY_OPERATIONAL_BALANCE_SUMMARY_20171129.xlsx"
daily_operational_balance_gtrm_input_file = "Z:/Charles/ORM/Data/DAILY_OPERATIONAL_BALANCE_SUMMARY_GTRM.xlsx"
date_mapping_file = "Z:/Charles/ORM/Data/Dates_Mapping.xlsx"

################# Load Libraries #################
import os
import pandas as pd
import time
import numpy as np
import matplotlib.pyplot as plt
plt.rcParams.update({'figure.max_open_warning': 0})
from matplotlib.ticker import FuncFormatter
import datetime
import dateutil
import subprocess
import imp
from scipy import stats
import gc
import pylab
# Load local functions
os.chdir(code_dir)
import waterfall
imp.reload(waterfall)
from waterfall import Waterfall
import UpdateDatabase_DB
imp.reload(UpdateDatabase_DB)
from UpdateDatabase_DB import read_db
import Functions_Analysis
imp.reload(Functions_Analysis)
from Functions_Analysis import summarizeTableByAttribute, summarizeTable, process_dda_GTRM, formatDate, render_mpl_table, color_function, color_function_client #, summarizeTableRatioOfTwoFactor, listisin

#################### Functions  ####################
def aggregate_summary_table_with_percentiles(summmary_result, attributes = None, reference_dates = None,digit=1,add_total=False):
#    summmary_result = result_total
#    summary_result = result_ope/result_total*1e9
    summarypage1 = summmary_result.copy()
    if attributes is None:
        attributes = summmary_result.columns
    if reference_dates is None:
#        current_date = summarypage1.head(1)
#        summarypage1 = summarypage1.head(2).transpose()
#        summarypage1.columns = [x.strftime("%Y-%m-%d") for x in summarypage1.columns]
#        summarypage1['DailyChg']=summarypage1.iloc[:,0]-summarypage1.iloc[:,1]
        
        current_date = summarypage1.loc[[today],:]
        summarypage1 = summarypage1.loc[[today,yesterday],:].transpose()
        summarypage1.columns = [x.strftime("%Y-%m-%d") for x in summarypage1.columns]
        summarypage1['change']=summarypage1.iloc[:,0]-summarypage1.iloc[:,1]
    else:
        current_date = summarypage1[summarypage1.index == np.datetime64(reference_dates[0])]
        summarypage1 = summarypage1[summarypage1.index.isin([np.datetime64(x) for x in reference_dates])].transpose()
        summarypage1.columns = [x.strftime("%Y-%m-%d") for x in summarypage1.columns]

    summarypage1 = round(summarypage1*1e-9,digit)
    summarypage1['Average']=round(summmary_result.apply(np.mean,axis=0)*1e-9,digit)
    summarypage1['Max']=round(summmary_result.apply(max,axis=0)*1e-9,digit)
    summarypage1['Min']=round(summmary_result.apply(min,axis=0)*1e-9,digit)
    summarypage1['Std']=round(summmary_result.apply(np.std,axis=0)*1e-9,digit)
    summarypage1 = summarypage1[summarypage1.index.isin(attributes)]
    
    percentiles = []
    for i in range(len(attributes)):
        p=stats.percentileofscore(summmary_result[attributes[i]], current_date[attributes[i]].values[0], 'rank')/100
        p = '{:.0%}'.format(p)                    
        percentiles.append(p)
    
    summarypage1['Percentile']=percentiles
    summarypage1.index = [x.replace("_"," ") if x != "_Missing" else x for x in summarypage1.index]
    return(summarypage1)

def aggregate_summary_table_with_percentiles_and_changes(summmary_result, attributes=None,digit=1):
#    summmary_result=result_total
    # statistics of changes
    if attributes is None:
        attributes=summmary_result.columns
    summarypage2 = summmary_result.copy()
    daycount = np.busday_count(np.datetime64(today).astype('datetime64[D]'), np.datetime64(yesterday).astype('datetime64[D]'))

#    today.astype('datetime64[D]')
#    np.datetime64(today)
#    daycount = np.busday_count(np.datetime64(today), np.datetime64(yesterday))
    
    current_diff = summarypage2.loc[[today,yesterday],:].diff(-1).dropna()
#    current_diff = summarypage2.diff(-1).head(1).dropna()
    summarypage2 = current_diff.transpose()
    summarypage2.columns = ["Change"]
    summarypage2 = round(summarypage2*1e-9,digit)
    summarypage2['Average']=round(summmary_result.diff(daycount).dropna().apply(np.nanmean,axis=0)*1e-9,digit)
    summarypage2['Max']=round(summmary_result.diff(daycount).dropna().apply(np.max,axis=0)*1e-9,digit)
    summarypage2['Min']=round(summmary_result.diff(daycount).dropna().apply(np.min,axis=0)*1e-9,digit)
    summarypage2['Std']=round(summmary_result.diff(daycount).dropna().apply(np.std,axis=0)*1e-9,digit)
    summarypage2 = summarypage2[summarypage2.index.isin(attributes)]
    percentiles = []
    for i in range(len(attributes)):
#        i = 0
        p = stats.percentileofscore(summmary_result.diff(daycount).dropna()[attributes[i]], current_diff[attributes[i]].values[0], kind = 'rank')/100
        p = '{:.0%}'.format(p)  
        percentiles.append(p)
        
    summarypage2['Percentile'] = percentiles
    summarypage2.index = [x.replace("_"," ") if x is not "_Missing" else x for x in summarypage2.index]
    
    return(summarypage2)

#################### Configuration  ####################
#Reporting Parameters
Reporting_Total_Deposit = 0
Reporting_Ope_Deposit = 0
Reporting_Total_Deposit_Prior_Day = 0
Reporting_Ope_Deposit_Prior_Day = 0

#Setup parameters
Parameters = pd.ExcelFile(input_file).parse("Sheet1")

if TEST_ENVIRONMENT:
#    end_date = datetime.datetime.strptime('2018-05-09 00:00:00', '%Y-%m-%d %H:%M:%S')     # historical data end date - later than benchmark date
    end_date = datetime.datetime.today()
    start_date = end_date - datetime.timedelta(days=5)
    benchmarkdate = [datetime.datetime.strptime('2017-11-10 00:00:00', '%Y-%m-%d %H:%M:%S')]
#    start_date = datetime.datetime.strptime('2017-12-27 00:00:00', '%Y-%m-%d %H:%M:%S')   # historical data start date - before benchmark date
#    benchmarkdate = [datetime.datetime.strptime('2017-11-07 00:00:00', '%Y-%m-%d %H:%M:%S'),\
#    datetime.datetime.strptime('2017-11-08 00:00:00', '%Y-%m-%d %H:%M:%S'),\
#    datetime.datetime.strptime('2017-11-09 00:00:00', '%Y-%m-%d %H:%M:%S'),\
#    datetime.datetime.strptime('2017-11-10 00:00:00', '%Y-%m-%d %H:%M:%S'),\
#    datetime.datetime.strptime('2017-11-13 00:00:00', '%Y-%m-%d %H:%M:%S'),\
#    datetime.datetime.strptime('2017-11-14 00:00:00', '%Y-%m-%d %H:%M:%S')]
else:
    end_date = max(read_db("Total_Deposits_By_BehaviorGroup").columns)
#    end_date = datetime.datetime.strptime('2018-09-05 00:00:00', '%Y-%m-%d %H:%M:%S')
    start_date = end_date - datetime.timedelta(days=5)
    benchmarkdate = [datetime.datetime.strptime('2017-11-10 00:00:00', '%Y-%m-%d %H:%M:%S')]

timewindow = [(end_date - datetime.timedelta(days=x)).strftime("%Y%m%d") for x in range((end_date-start_date).days+1)]
benchmarkwindow = [x.strftime("%Y%m%d") for x in benchmarkdate]
timewindow = sorted(list(set(benchmarkwindow+timewindow)))

date_newmodel = datetime.datetime.strptime("2017-4-26 00:00:00", '%Y-%m-%d %H:%M:%S')
# create folder hierarchy
if not os.path.exists(output_dir):
    os.makedirs(output_dir)
if not os.path.exists(weekly_table_dir):
    os.makedirs(weekly_table_dir)
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
plot_folder_weekly = weekly_table_dir
if not os.path.exists(plot_folder_weekly):
    os.mkdir(plot_folder_weekly)

# formats and styles
pd.options.mode.chained_assignment = None
plt.style.use('ggplot')
styles = ['b-','g-','r-','c-','m-','y-','k-']
styles2 = ['b--','g--','r--','c--','m--','y--','k--']

def percentages(x):
    return '{:.1%}'.format(x)
def millions(x, pos=0):
    return '{:,.1f}'.format(x*1e-6)
def billions(x, pos=0):
    return '{:,.1f}B'.format(x*1e-9)
def billionsinteger(x, pos=0):
    return '{:,.0f}B'.format(x*1e-9)

formatter_percentages = FuncFormatter('{0:.0%}'.format)
formatter_millions = FuncFormatter(millions)
formatter_billions = FuncFormatter(billions)
formatter_billionsinteger = FuncFormatter(billionsinteger)

#################### Finish Configuration ####################

start_time = time.time()
#################### Read and Processing Data #########
table_name = "OPS_DEPOSIT_PBUP_FBO_DDA_GTRM"
data_all = read_db(table_name,timewindow)
data_all = process_dda_GTRM(data_all)

#new column names
column_name_operational = "Operational_Deposits"
column_name_spot = "Total_Deposits"
column_name_excess = "Excess_Deposits"
attribute_aggs = [column_name_spot,column_name_operational,column_name_excess]
column_name_local = "PRINCIPAL_BAL"
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
column_name_product_typenew = "ProductTypeNew"
column_name_GL_type = "GLProdType"

column_name_fx = "FX_BASE_TO_USD"

dates = data_all["AS_OF_DATE"].drop_duplicates().sort_values(ascending=False)
today = dates[0]
yesterday = dates[1]
#yesterday = today + dateutil.relativedelta.relativedelta(weeks=-2)
#yesterday = today + dateutil.relativedelta.relativedelta(days=-8)

lastweek = today + dateutil.relativedelta.relativedelta(weeks=-1)

if not any(lastweek == dates):
    while True:
        #a week earlier may be a holidy. So search the nearest day before that day
        tmp = read_db(table_name,[lastweek.strftime("%Y%m%d")])
        if len(tmp) >0:
            break
        else:
            lastweek = lastweek + dateutil.relativedelta.relativedelta(days=-1)
    data_all = data_all.append(process_dda_GTRM(tmp))
    dates = data_all["AS_OF_DATE"].drop_duplicates().sort_values(ascending=False)
if not any(yesterday == dates):
    while True:
        #a week earlier may be a holidy. So search the nearest day before that day
        tmp = read_db(table_name,[yesterday.strftime("%Y%m%d")])
        if len(tmp) >0:
            break
        else:
            yesterday = yesterday + dateutil.relativedelta.relativedelta(days=-1)
    data_all = data_all.append(process_dda_GTRM(tmp))
    dates = data_all["AS_OF_DATE"].drop_duplicates().sort_values(ascending=False)
    
if not pd.isnull(Parameters.loc['CurrentDate',"Date"]):
    currentsnapshot = Parameters.loc['CurrentDate',"Date"]
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
result = summarizeTableByAttribute(data_all,attribute_aggs)
if os.path.exists(daily_operational_balance_gtrm_input_file):
    daily_balance_gtrm = pd.ExcelFile(daily_operational_balance_gtrm_input_file).parse('Sheet1',index_col=0)
    result_update = result.copy()
    result_update['UPDATED_ON'] = datetime.datetime.today()
    daily_balance_gtrm = daily_balance_gtrm[~daily_balance_gtrm.index.isin(result_update.index)].append(result_update).sort_index()
else:
    daily_balance = pd.ExcelFile(daily_operational_balance_input_file).parse('Sheet1')
    daily_balance.rename(columns={'SPOT BALANCE':column_name_spot,'OPERATIONAL BALANCE':column_name_operational,'EXCESS BALANCE':column_name_excess},inplace=True)
    daily_balance.set_index('AS_OF_DATE',inplace=True)
    daily_balance = daily_balance[[column_name_spot,column_name_operational,column_name_excess]]
    daily_balance_gtrm = daily_balance[~daily_balance.index.isin(result.index)].append(result).sort_index()
    daily_balance_gtrm['UPDATED_ON'] = datetime.datetime.today()
daily_balance_gtrm.to_excel(daily_operational_balance_gtrm_input_file)

daily_balance_gtrm = daily_balance_gtrm.dropna()
daily_balance_gtrm = daily_balance_gtrm[daily_balance_gtrm.index<=today]

Reporting_Total_Deposit = daily_balance_gtrm.loc[today,column_name_spot]
Reporting_Ope_Deposit = daily_balance_gtrm.loc[today,column_name_operational]
Reporting_Total_Deposit_Prior_Day = daily_balance_gtrm.loc[yesterday,column_name_spot]
Reporting_Ope_Deposit_Prior_Day = daily_balance_gtrm.loc[yesterday,column_name_operational]

while True:
    try:
        plt.figure()
        daily_balance_gtrm_figure = daily_balance_gtrm.copy()
        daily_balance_gtrm_figure = daily_balance_gtrm_figure.loc[:,daily_balance_gtrm_figure.columns!="UPDATED_ON"]
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
        
        daily_balance_gtrm_figure.to_excel(plot_folder+"/Summary_DepositChartLong.xlsx")
        break
    except:
        pass

#summary3
summmary_result = daily_balance_gtrm.sort_index(ascending=False)[[column_name_spot,column_name_operational,column_name_excess]]
aggregate_summary_table_with_percentiles(summmary_result).to_latex(plot_folder+"/Summary_DepositTable.tex",longtable=True)
aggregate_summary_table_with_percentiles(summmary_result,digit=5).to_excel(plot_folder+"/Summary_DepositTable.xlsx")

aggregate_summary_table_with_percentiles(summmary_result,reference_dates = [today, yesterday, lastweek]).to_latex(plot_folder+"/Summary_Weekly_DepositTable.tex",longtable=True)
aggregate_summary_table_with_percentiles(summmary_result,reference_dates = [today, yesterday, lastweek],digit=5).to_excel(plot_folder+"/Summary_Weekly_DepositTable.xlsx")

summmary_result = daily_balance_gtrm.sort_index(ascending=False)[[column_name_spot,column_name_operational,column_name_excess]]
summmary_result = summmary_result[summmary_result.index>=date_newmodel]
aggregate_summary_table_with_percentiles(summmary_result).to_latex(plot_folder+"/Summary_DepositTable_NewModel.tex",longtable=True)
aggregate_summary_table_with_percentiles(summmary_result,digit=5).to_excel(plot_folder+"/Summary_DepositTable_NewModel.xlsx")

aggregate_summary_table_with_percentiles(summmary_result,reference_dates = [today, yesterday, lastweek]).to_latex(plot_folder+"/Summary_Weekly_DepositTable_NewModel.tex",longtable=True)
aggregate_summary_table_with_percentiles(summmary_result,reference_dates = [today, yesterday, lastweek],digit=5).to_excel(plot_folder+"/Summary_Weekly_DepositTable_NewModel.xlsx")

summmary_result = daily_balance_gtrm.sort_index(ascending=False)[[column_name_spot,column_name_operational,column_name_excess]]
aggregate_summary_table_with_percentiles_and_changes(summmary_result).to_latex(plot_folder+"/Summary_DepositChgTable.tex",longtable=True)
aggregate_summary_table_with_percentiles_and_changes(summmary_result,digit=5).to_excel(plot_folder+"/Summary_DepositChgTable.xlsx")

summmary_result = daily_balance_gtrm.sort_index(ascending=False)[[column_name_spot,column_name_operational,column_name_excess]]
summmary_result = summmary_result[summmary_result.index>=date_newmodel]
aggregate_summary_table_with_percentiles_and_changes(summmary_result).to_latex(plot_folder+"/Summary_DepositChgTable_NewModel.tex",longtable=True)
aggregate_summary_table_with_percentiles_and_changes(summmary_result,digit=5).to_excel(plot_folder+"/Summary_DepositChgTable_NewModel.xlsx")

daily_balance=[]
daily_balance_gtrm=[]
result_tex=[]
figure=[]
gc.collect()

##Decompose for Summary
# describe hedge fund excluded
def aggregate_summary_table_format_combined(data_all,currentsnapshot,benchmarkdate,attribute_groupby,attribute_agg,need_aggregate=False,group_agg=None,group_name=None,top_number=7):
#    data_all = data_all_tmp
#    attribute_groupby = attribute_groupby
#    attribute_agg = column_name_spot
#    group_agg = Alternatives
#    currentsnapshot = today
#    benchmarkdate = yesterday

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
#    currentdate = max(currentsnapshot_tmp)
#    numdaysago = 30
    result_d = result.columns - (currentdate-datetime.timedelta(days=numdaysago))
    if any(result_d.days<=0):
        result_d = abs(result_d[result_d.days<=0])
        
    result_d = result_d==min(result_d)
    column = [i for i, x in enumerate(result_d) if x]
    return(result.columns[column][0])

def concat_history(column_name_spot,column_name_group,result_total):
    # get the saved historical data
#    column_name_spot = attribute_agg
#    column_name_group = attribute_groupby
#    result_total = table_orig.transpose()
#    if column_name_group == 'ParentCompy':
#        if column_name_spot == "Total_Deposits":
#            file_name = "TotalDepositTableByParent"
#        elif column_name_spot == "Operational_Deposits":
#            file_name = "OperationalDepositTableByParent"
#    else:
    file_name = column_name_spot+"_By_"+column_name_group
    result_history = read_db(file_name).transpose()
    result_history['Total'] = result_history.apply(np.nansum,1)
    if FED_ENVIRONMENT:
        if 'Funds Miss Model Info' in result_history.columns:
            result_history['Excluded Funds'] = result_history['Excluded Funds'] + result_history['Funds Miss Model Info']
            del result_history['Funds Miss Model Info']
        if 'Funds Miss Region Info' in result_history.columns:
            result_history['Excluded Funds'] = result_history['Excluded Funds'] + result_history['Funds Miss Region Info']
            del result_history['Funds Miss Region Info']

    if len(result_history.columns)!=len(result_total.columns):
        if all(result_total.columns.isin(result_history.columns)) and all(result_total.index.isin(result_history.index)):
            result_history = result_history
        else:
            print("Warning: concat history error.")
    else:
        if not (abs(result_history.loc[result_total.index,:]-result_total)<1e6).all().all():
            print("Warning: Numbers differ - " +column_name_spot+"_By_"+column_name_group)
        result_history = result_total[~result_total.index.isin(result_history.index)].append(result_history)
    
    result_total = result_history[result_history.index<=max(result_total.index)].sort_index(ascending=False)
    result_total = result_total.fillna(0)
    return(result_total)
    
def aggregate_summary_table(data_all,currentsnapshot,benchmarkdate,attribute_groupby,attribute_agg,need_aggregate=False,group_agg=None,group_name=None,top_number=7,sort_by="DiffToBck",average_period=30,digit=2):
#    data_all = data_all_tmp
#    attribute_groupby = attribute_groupby
#    attribute_agg = column_name_spot
#    currentsnapshot = today
#    benchmarkdate = yesterday
#    need_aggregate=need_aggregate,top_number=top_number,sort_by=sort_by

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
    table_orig = summarizeTable(data_date,attribute_groupby,attribute_agg).sort_values(currentsnapshot_tmp,ascending=False)
    table_orig = concat_history(attribute_agg,attribute_groupby,table_orig.transpose()).transpose()
    table_orig = table_orig[table_orig.index != "Total"]
    
    table_all = table_orig[currentsnapshot_tmp].apply(np.nanmean,axis=1).to_frame("AsOfPeriod")
    table_all["Benchmark"] = table_orig[benchmarkdate_tmp].apply(np.nanmean,axis=1)
#    sum(table_all.index=="Total")
    #get average of one month time window. The rolling windowns starts at the maximum of as of period and look back 30 days.
    month_dates = [max(currentsnapshot_tmp) - datetime.timedelta(days=x) for x in range(average_period)]
    result_one_month = table_orig[list(set(month_dates).intersection(set(table_orig.columns)))].apply(np.nanmean,axis=1)

    table_all["MonthAvg"] = result_one_month

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
        other = table_all.loc["UnregulatedFund",:]
        table_all = table_all[~table_all.index.isin(["UnregulatedFund"])]
        other.name = "UnregulatedFunds"
        table_all = table_all.append(other)

    if "_Missing" in table_all.index:
        other = table_all.loc["_Missing",:]
        table_all = table_all[~table_all.index.isin(["_Missing"])]
        other.name = "_Missing"
        table_all = table_all.append(other)
        
    if "Other" in table_all.index:
        other = table_all.loc["Other",:]
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

    table_all.iloc[:,0:5]=round(table_all.iloc[:,0:5]/1e9,digit)
    return(table_all[["AsOfPeriod","Benchmark","DiffToBck","%DiffToBck","MonthAvg","DiffToAvg","%DiffToAvg"]])

#pivot tables
#pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))))
#pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute,output_excel=False,output_tex=False)
#result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute,output_excel=False,output_tex=False,color_function=color_function,client_subset=top_clients)
#pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))),color_function=color_function)
def pivot_tables(data_all_tmp, index, pivot_attributes, today=today,yesterday=yesterday,func = np.nansum, color_function=None, output_excel = True, output_tex = True):
    tmp_data_all = data_all_tmp[data_all_tmp["AS_OF_DATE"].isin([today,yesterday])]
#    tmp_data_all.loc[[x not in ["SSBT USA","SSBT GmbH","SSBT London"] for x in tmp_data_all[column_name_company]],column_name_company] = "Others"
    tmp_data_all.loc[[x not in ["DDA-DOMESTIC","IBDDA-DOMESTIC","IBDDA-EUROPEAN",'IBDDA-CAYMAN','IBDDA-LONDON'] for x in tmp_data_all['ProductTypes']],'ProductTypes'] = "Others"
    
    if index == column_name_parent or index == column_name_group:
        tmp_data_all.loc[[x not in ["SSBT USA","SSBT GmbH","SSBT London"] for x in tmp_data_all[column_name_company]],column_name_company] = "Others"
#        tmp_data_all.loc[[x not in ["USD","EUR","GBP"] for x in tmp_data_all[column_name_currency]],column_name_currency] = "Others"

    for j in range(len(pivot_attributes)):
        data_all_today = tmp_data_all[tmp_data_all["AS_OF_DATE"]==today]
        data_all_yesterday = tmp_data_all[tmp_data_all["AS_OF_DATE"]==yesterday]
        data_all_today[pivot_attributes[j]] = data_all_today[pivot_attributes[j]].fillna("Missing")
        data_all_yesterday[pivot_attributes[j]] = data_all_yesterday[pivot_attributes[j]].fillna("Missing")

        pivot_today = pd.pivot_table(data_all_today, values=column_name_spot, index=attribute_groupby, columns=pivot_attributes[j], aggfunc=func, fill_value=None, margins=False, dropna=True, margins_name='All')
        pivot_yesterday = pd.pivot_table(data_all_yesterday, values=column_name_spot, index=attribute_groupby, columns=pivot_attributes[j], aggfunc=func, fill_value=None, margins=False, dropna=True, margins_name='All')
#        summarypivot = pivot_today - pivot_yesterday
        summarypivot = pivot_today.fillna(0) - pivot_yesterday.fillna(0)

        if 'Others' in summarypivot.columns:
            summarypivot.insert(len(summarypivot.columns)-1,'Others',summarypivot.pop('Others'))
        if 'Others' in pivot_today.columns:
            pivot_today.insert(len(pivot_today.columns)-1,'Others',pivot_today.pop('Others'))
        summarypivot.loc['Sum',:]=summarypivot.apply(np.nansum,axis=0)
        summarypivot['Sum']=summarypivot.apply(np.nansum,axis=1)
        pivot_today.loc['Sum',:]=pivot_today.apply(np.nansum,axis=0)
        pivot_today['Sum']=pivot_today.apply(np.nansum,axis=1)
        
        # table 3: compare to history
#        pivot_today = pivot_today[pivot_today.index!='Sum']
#        pivot_today = pivot_today.loc[:,pivot_today.columns!='Sum']
        pivot_today = pivot_today.fillna(0)
        pivot_yesterday = pivot_yesterday.fillna(0)
        
        pivot_diff = pivot_today.copy()
        pivot_percentile = pivot_today.copy()
        table_name = "Total_Deposits_By_"+index+"_and_"+pivot_attributes[j]
        
        if color_function is not None:
            history = read_db(table_name)
            if history.__class__ == pd.DataFrame:
                table_name = "Total_Deposits_By_"+pivot_attributes[j]
                history_sum_attribute = read_db(table_name)
                table_name = "Total_Deposits_By_"+index
                history_sum_index = read_db(table_name)
    
                for k in range(pivot_today.shape[0]):
    #                k = 0
                    for l in range(pivot_today.shape[1]):
    #                    l = 1
                        index_category = pivot_today.index[k]
                        attribute_category = pivot_today.columns[l]
                        if (index_category,attribute_category) in history.index:
                            history_item = history.loc[index_category,attribute_category]
                            history_item = history_item[(history_item.index <= today) & (history_item.index >= today - datetime.timedelta(days=365))]
                            pivot_diff.loc[index_category,attribute_category] = pivot_today.iloc[k,l] - np.nanmean(history_item)
                            pivot_percentile.loc[index_category,attribute_category] = stats.percentileofscore(history_item, pivot_today.iloc[k,l], 'rank')/100
                        elif index_category == "Sum" and attribute_category != "Sum":
                            history_item = history_sum_attribute.loc[attribute_category,:]
                            pivot_diff.loc[index_category,attribute_category] = pivot_today.iloc[k,l] - np.nanmean(history_item)
                            pivot_percentile.loc[index_category,attribute_category] = stats.percentileofscore(history_item, pivot_today.iloc[k,l], 'rank')/100
                        elif index_category != "Sum" and attribute_category == "Sum":
                            if index_category in history_sum_index.index:
                                history_item = history_sum_index.loc[index_category,:]
                                pivot_diff.loc[index_category,attribute_category] = pivot_today.iloc[k,l] - np.nanmean(history_item)
                                pivot_percentile.loc[index_category,attribute_category] = stats.percentileofscore(history_item, pivot_today.iloc[k,l], 'rank')/100
                            else:
                                pivot_diff.loc[index_category,attribute_category] = np.nan
                                pivot_percentile.loc[index_category,attribute_category] = np.nan
                        elif index_category == "Sum" and attribute_category == "Sum":
                            history_item = history_sum_index.apply(np.nansum,axis=0)
                            pivot_diff.loc[index_category,attribute_category] = pivot_today.iloc[k,l] - np.nanmean(history_item)
                            pivot_percentile.loc[index_category,attribute_category] = stats.percentileofscore(history_item, pivot_today.iloc[k,l], 'rank')/100
                        else:
                            pivot_diff.loc[index_category,attribute_category] = np.nan
                            pivot_percentile.loc[index_category,attribute_category] = np.nan
                tmp_results = round(pivot_diff/1e9,2).applymap(str)+"/"+pivot_percentile.applymap(percentages)
                tmp_results.replace("nan/nan%","",inplace=True)
                if pivot_attributes[j] == column_name_group:
                    tmp_results = tmp_results.rename(columns={"Funds Miss Model Info":"Miss Info","Multi-Day High Clients":"MDH Clients","Multi-Day Low Clients":"MDL Clients","Non-Discretionary Funds":"ND Funds"})
#                client_subset=top_clients
                fig = render_mpl_table(tmp_results, data_color = pivot_diff.fillna(0), func=color_function, col_width=2.0,row_height=0.4)
                fig.savefig(plot_folder+"/Summary_DailyValue_Diff_"+attribute_groupby+"_by_"+pivot_attributes[j]+".png",bbox_inches='tight')
                
        #Output tex and excel format for table 1 and table 2
        if pivot_attributes[j] == column_name_group:
            summarypivot = summarypivot.rename(columns={"Funds Miss Model Info":"Miss Info","Multi-Day High Clients":"MDH Clients","Multi-Day Low Clients":"MDL Clients","Non-Discretionary Funds":"ND Funds"})
            pivot_today = pivot_today.rename(columns={"Funds Miss Model Info":"Miss Info","Multi-Day High Clients":"MDH Clients","Multi-Day Low Clients":"MDL Clients","Non-Discretionary Funds":"ND Funds"})
        if output_tex:
            round(summarypivot/1e9,2).fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attributes[j]+".tex",column_format='l'+'r'*len(summarypivot.columns),longtable=True)
            round(pivot_today/1e9,2).fillna("").to_latex(plot_folder+"/Summary_DailyValue_"+attribute_groupby+"_by_"+pivot_attributes[j]+".tex",column_format='l'+'r'*len(summarypivot.columns),longtable=True)
        if output_excel:
            summarypivot.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attributes[j]+".xlsx")
            pivot_today.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attributes[j]+".xlsx")
        if j == 0:
            pivots = summarypivot
        else:
            pivots = pd.concat([pivots,summarypivot],axis=1)
    return(pivots)

def pivot_tables_client(data_all_tmp, index=column_name_parent, pivot_attributes=[column_name_currency], today=today,yesterday=yesterday,func = np.nansum, color_function=color_function_client):
    tmp_data_all = data_all_tmp[data_all_tmp["AS_OF_DATE"].isin([today,yesterday])]
#    tmp_data_all.loc[[x not in ["SSBT USA","SSBT GmbH","SSBT London"] for x in tmp_data_all[column_name_company]],column_name_company] = "Others"
    tmp_data_all.loc[[x not in ["DDA-DOMESTIC","IBDDA-DOMESTIC","IBDDA-EUROPEAN",'IBDDA-CAYMAN','IBDDA-LONDON'] for x in tmp_data_all['ProductTypes']],'ProductTypes'] = "Others"
    
    if index == column_name_parent or index == column_name_group:
        tmp_data_all.loc[[x not in ["SSBT USA","SSBT GmbH","SSBT London"] for x in tmp_data_all[column_name_company]],column_name_company] = "Others"
#        tmp_data_all.loc[[x not in ["USD","EUR","GBP"] for x in tmp_data_all[column_name_currency]],column_name_currency] = "Others"

    j = 0
    data_all_today = tmp_data_all[tmp_data_all["AS_OF_DATE"]==today]
    data_all_yesterday = tmp_data_all[tmp_data_all["AS_OF_DATE"]==yesterday]
    data_all_today[pivot_attributes[j]] = data_all_today[pivot_attributes[j]].fillna("Missing")
    data_all_yesterday[pivot_attributes[j]] = data_all_yesterday[pivot_attributes[j]].fillna("Missing")

    pivot_today = pd.pivot_table(data_all_today, values=column_name_spot, index=attribute_groupby, columns=pivot_attributes[j], aggfunc=func, fill_value=None, margins=False, dropna=True, margins_name='All')
    pivot_yesterday = pd.pivot_table(data_all_yesterday, values=column_name_spot, index=attribute_groupby, columns=pivot_attributes[j], aggfunc=func, fill_value=None, margins=False, dropna=True, margins_name='All')
#        summarypivot = pivot_today - pivot_yesterday
    summarypivot = pivot_today.fillna(0) - pivot_yesterday.fillna(0)

    if 'Others' in summarypivot.columns:
        summarypivot.insert(len(summarypivot.columns)-1,'Others',summarypivot.pop('Others'))
    if 'Others' in pivot_today.columns:
        pivot_today.insert(len(pivot_today.columns)-1,'Others',pivot_today.pop('Others'))
    summarypivot.loc['Sum',:]=summarypivot.apply(np.nansum,axis=0)
    summarypivot['Sum']=summarypivot.apply(np.nansum,axis=1)
    pivot_today.loc['Sum',:]=pivot_today.apply(np.nansum,axis=0)
    pivot_today['Sum']=pivot_today.apply(np.nansum,axis=1)
    
    # table 3: compare to history
#        pivot_today = pivot_today[pivot_today.index!='Sum']
#        pivot_today = pivot_today.loc[:,pivot_today.columns!='Sum']
    pivot_today = pivot_today.fillna(0)
    pivot_yesterday = pivot_yesterday.fillna(0)
    
    pivot_diff = pivot_today.copy()
    pivot_percentile = pivot_today.copy()
    table_name = "Total_Deposits_By_"+index+"_and_"+pivot_attributes[j]
    
    history = read_db(table_name)
    if history.__class__ == pd.DataFrame:
        table_name = "Total_Deposits_By_"+pivot_attributes[j]
        history_sum_attribute = read_db(table_name)
        table_name = "Total_Deposits_By_"+index
        history_sum_index = read_db(table_name)

        for k in range(pivot_today.shape[0]):
#                k = 0
            for l in range(pivot_today.shape[1]):
#                    l = 1
                index_category = pivot_today.index[k]
                attribute_category = pivot_today.columns[l]
                if (index_category,attribute_category) in history.index:
                    history_item = history.loc[index_category,attribute_category]
                    history_item = history_item[(history_item.index <= today) & (history_item.index >= today - datetime.timedelta(days=365))]
                    pivot_diff.loc[index_category,attribute_category] = pivot_today.iloc[k,l] - np.nanmean(history_item)
                    pivot_percentile.loc[index_category,attribute_category] = stats.percentileofscore(history_item, pivot_today.iloc[k,l], 'rank')/100
                elif index_category == "Sum" and attribute_category != "Sum":
                    history_item = history_sum_attribute.loc[attribute_category,:]
                    pivot_diff.loc[index_category,attribute_category] = pivot_today.iloc[k,l] - np.nanmean(history_item)
                    pivot_percentile.loc[index_category,attribute_category] = stats.percentileofscore(history_item, pivot_today.iloc[k,l], 'rank')/100
                elif index_category != "Sum" and attribute_category == "Sum":
                    if index_category in history_sum_index.index:
                        history_item = history_sum_index.loc[index_category,:]
                        pivot_diff.loc[index_category,attribute_category] = pivot_today.iloc[k,l] - np.nanmean(history_item)
                        pivot_percentile.loc[index_category,attribute_category] = stats.percentileofscore(history_item, pivot_today.iloc[k,l], 'rank')/100
                    else:
                        pivot_diff.loc[index_category,attribute_category] = np.nan
                        pivot_percentile.loc[index_category,attribute_category] = np.nan
                elif index_category == "Sum" and attribute_category == "Sum":
                    history_item = history_sum_index.apply(np.nansum,axis=0)
                    pivot_diff.loc[index_category,attribute_category] = pivot_today.iloc[k,l] - np.nanmean(history_item)
                    pivot_percentile.loc[index_category,attribute_category] = stats.percentileofscore(history_item, pivot_today.iloc[k,l], 'rank')/100
                else:
                    pivot_diff.loc[index_category,attribute_category] = np.nan
                    pivot_percentile.loc[index_category,attribute_category] = np.nan
        tmp_results = round(pivot_diff/1e9,2).applymap(str)+"/"+pivot_percentile.applymap(percentages)
        tmp_results.replace("nan/nan%","",inplace=True)
        if pivot_attributes[j] == column_name_group:
            tmp_results = tmp_results.rename(columns={"Funds Miss Model Info":"Miss Info","Multi-Day High Clients":"MDH Clients","Multi-Day Low Clients":"MDL Clients","Non-Discretionary Funds":"ND Funds"})
#                client_subset = top_clients
    else:
        tmp_results = []    
    return(tmp_results,pivot_diff)


###change by attributes
#data prep
pivot_attributes = [column_name_group,column_name_region,column_name_company,'ProductTypes',column_name_currency,column_name_division,column_name_stylegroup,column_name_market]
data_all_tmp = data_all[attribute_aggs +["AS_OF_DATE","ND_IND",column_name_parent]+pivot_attributes]

##behavior groups
attribute_groupby = column_name_group
result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
if FED_ENVIRONMENT:
    result = result.reindex(["Excluded Funds","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
else:    
    result = result.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
result.columns.name = "Fund/Client Type"
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
result_total['Total'] = result_total.apply(np.nansum,axis=1)
result_total = concat_history(column_name_spot,attribute_groupby,result_total)
summarypage1 = aggregate_summary_table_with_percentiles(result_total)
if FED_ENVIRONMENT:
    summarypage1 = summarypage1.reindex(["Excluded Funds","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
else:
    summarypage1 = summarypage1.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
if FED_ENVIRONMENT:
    summarypage1diff = summarypage1diff.reindex(["Excluded Funds","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
else:
    summarypage1diff = summarypage1diff.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")

#attrition analysis
df = result_total.loc[today,:] - result_total.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_total.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 7), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_spot, outfile=plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".png")

result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
result_ope = concat_history(column_name_operational,attribute_groupby,result_ope)
summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
if FED_ENVIRONMENT:
    summarypage1 = summarypage1.reindex(["Excluded Funds","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
else:
    summarypage1 = summarypage1.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
if FED_ENVIRONMENT:
    summarypage1diff = summarypage1diff.reindex(["Excluded Funds","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
else:
    summarypage1diff = summarypage1diff.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")
#attrition analysis
df = result_ope.loc[today,:] - result_ope.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_ope.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 7), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_operational, outfile=plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".png")

summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
if FED_ENVIRONMENT:
    summarypage1 = summarypage1.reindex(["Excluded Funds","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
else:
    summarypage1 = summarypage1.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
summarypage1.iloc[:,:7] = summarypage1.iloc[:,:7].applymap(percentages)
summarypage1.to_latex(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".xlsx")

#weekly dashboard
result_total_today = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total_today.columns = [today.strftime("%m/%d/%y"),yesterday.strftime("%m/%d/%y"),"change"]
result_total_lastweek = aggregate_summary_table_format_combined(data_all_tmp,today,lastweek,attribute_groupby,column_name_spot)
result_total_lastweek.columns = [today.strftime("%m/%d/%y"),lastweek.strftime("%m/%d/%y"),"change"]
result_total = pd.concat([result_total_today.iloc[:,0:2],result_total_lastweek.iloc[:,1:2]],axis=1)
                        
result_operational_today = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational_today.columns = [today.strftime("%m/%d/%y"),yesterday.strftime("%m/%d/%y"),"change"]
result_operational_lastweek = aggregate_summary_table_format_combined(data_all_tmp,today,lastweek,attribute_groupby,column_name_operational)
result_operational_lastweek.columns = [today.strftime("%m/%d/%y"),lastweek.strftime("%m/%d/%y"),"change"]
result_operational = pd.concat([result_operational_today.iloc[:,0:2],result_operational_lastweek.iloc[:,1:2]],axis=1)

result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,2)
#result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
result = result.reindex(["Excluded Funds","Funds Miss Model Info","Non-Discretionary Funds","Intraday Clients","Multi-Day Low Clients","Multi-Day High Clients","Total"])
result.columns.name = "Fund/Client Type"

result.to_latex(plot_folder+"/Summary_Weekly_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Weekly_"+attribute_groupby+".xlsx")
result.to_excel(plot_folder_weekly+"/Summary_Weekly_"+attribute_groupby+".xlsx")

#benchmark dates
result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

pivot_tables(data_all_tmp, index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))),color_function=color_function)

### Region
attribute_groupby = column_name_region

result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
if FED_ENVIRONMENT:
    result = result.reindex(["Excluded Funds","AMERICAS","EMEA","APAC","UNKNOWN","Total"])
else:
    result = result.reindex(["Excluded Funds","Funds Miss Region Info","AMERICAS","EMEA","APAC","UNKNOWN","Total"])
result.columns.name = "Region"

result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
result_total['Total'] = result_total.apply(np.nansum,axis=1)
result_total = concat_history(column_name_spot,attribute_groupby,result_total)
summarypage1 = aggregate_summary_table_with_percentiles(result_total)
if FED_ENVIRONMENT:
    summarypage1 = summarypage1.reindex(["Excluded Funds","AMERICAS","EMEA","APAC","UNKNOWN","Total"])
else:
    summarypage1 = summarypage1.reindex(["Excluded Funds","Funds Miss Region Info","AMERICAS","EMEA","APAC","UNKNOWN","Total"])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
if FED_ENVIRONMENT:
    summarypage1diff = summarypage1diff.reindex(["Excluded Funds","AMERICAS","EMEA","APAC","UNKNOWN","Total"])
else:
    summarypage1diff = summarypage1diff.reindex(["Excluded Funds","Funds Miss Region Info","AMERICAS","EMEA","APAC","UNKNOWN","Total"])
    
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")

#attrition analysis
df = result_total.loc[today,:] - result_total.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_total.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_spot, outfile=plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".png")

result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
result_ope = concat_history(column_name_operational,attribute_groupby,result_ope)
summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
if FED_ENVIRONMENT:
    summarypage1 = summarypage1.reindex(["Excluded Funds","AMERICAS","EMEA","APAC","UNKNOWN","Total"])
else:
    summarypage1 = summarypage1.reindex(["Excluded Funds","Funds Miss Region Info","AMERICAS","EMEA","APAC","UNKNOWN","Total"])

summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
if FED_ENVIRONMENT:
    summarypage1diff = summarypage1diff.reindex(["Excluded Funds","AMERICAS","EMEA","APAC","UNKNOWN","Total"])
else:
    summarypage1diff = summarypage1diff.reindex(["Excluded Funds","Funds Miss Region Info","AMERICAS","EMEA","APAC","UNKNOWN","Total"])
    
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")

df = result_ope.loc[today,:] - result_ope.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_ope.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_operational, outfile=plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".png")

summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
if FED_ENVIRONMENT:
    summarypage1 = summarypage1.reindex(["Excluded Funds","AMERICAS","EMEA","APAC","UNKNOWN","Total"])
else:
    summarypage1 = summarypage1.reindex(["Excluded Funds","Funds Miss Region Info","AMERICAS","EMEA","APAC","UNKNOWN","Total"])
    
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

pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))),color_function=color_function)

### Legal Entity
attribute_groupby = column_name_company

result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
result = result.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
result.columns.name = attribute_groupby

result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
result_total['Total'] = result_total.apply(np.nansum,axis=1)
result_total = concat_history(column_name_spot,attribute_groupby,result_total)
summarypage1 = aggregate_summary_table_with_percentiles(result_total)
summarypage1 = summarypage1.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
summarypage1diff = summarypage1diff.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")
#attrition analysis
df = result_total.loc[today,:] - result_total.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_total.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_spot, outfile=plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".png")

result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
result_ope = concat_history(column_name_operational,attribute_groupby,result_ope)
summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
summarypage1 = summarypage1.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
summarypage1diff = summarypage1diff.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")
#attrition analysis
df = result_ope.loc[today,:] - result_ope.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_ope.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_operational, outfile=plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".png")

summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
summarypage1 = summarypage1.reindex(['SSBT USA','SSBT GmbH','SSBT London','SSBT Hong Kong','SSBT Canada', 'SSBT Sydney','SSBT Singapore','SSBT Jersey','Total'])
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
result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
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

pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))),color_function=color_function)

### Currency
attribute_groupby = column_name_currency
result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
result = result.reindex(['USD','EUR','GBP','JPY','AUD','CAD','CHF','Others','Total'])
result.columns.name = attribute_groupby
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
result_total['Total'] = result_total.apply(np.nansum,axis=1)
result_total = concat_history(column_name_spot,attribute_groupby,result_total)
summarypage1 = aggregate_summary_table_with_percentiles(result_total)
summarypage1 = summarypage1.reindex(['USD','EUR','GBP','JPY','AUD','CAD','CHF','Others','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
summarypage1diff = summarypage1diff.reindex(['USD','EUR','GBP','JPY','AUD','CAD','CHF','Others','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")

#attrition analysis
df = result_total.loc[today,:] - result_total.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_total.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_spot, outfile=plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".png")

result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
result_ope = concat_history(column_name_operational,attribute_groupby,result_ope)
summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
summarypage1 = summarypage1.reindex(['USD','EUR','GBP','JPY','AUD','CAD','CHF','Others','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
summarypage1diff = summarypage1diff.reindex(['USD','EUR','GBP','JPY','AUD','CAD','CHF','Others','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")
#attrition analysis
df = result_ope.loc[today,:] - result_ope.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_ope.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_operational, outfile=plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".png")

summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
summarypage1 = summarypage1.reindex(['USD','EUR','GBP','JPY','AUD','CAD','CHF','Others','Total'])
summarypage1.iloc[:,:7] = summarypage1.iloc[:,:7].applymap(percentages)
summarypage1.to_latex(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".xlsx")

result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))),color_function=color_function)

### Product Types
attribute_groupby = 'ProductTypes'
data_all_tmp_producttype = data_all_tmp[data_all_tmp[attribute_groupby].isin(['DDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-DOMESTIC'])]
result_total = aggregate_summary_table_format_combined(data_all_tmp_producttype,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp_producttype,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
result = result.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','Total'])
result.index = ['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','Total - SSBT USA']
result.columns.name = attribute_groupby
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
result_total['Total'] = result_total.apply(np.nansum,axis=1)
result_total = concat_history(column_name_spot,attribute_groupby,result_total)
summarypage1 = aggregate_summary_table_with_percentiles(result_total)
summarypage1 = summarypage1.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
summarypage1diff = summarypage1diff.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")
#attrition analysis
df = result_total.loc[today,:] - result_total.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_total.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_spot, outfile=plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".png")

result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
result_ope = concat_history(column_name_operational,attribute_groupby,result_ope)
summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
summarypage1 = summarypage1.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
summarypage1diff = summarypage1diff.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")
#attrition analysis
df = result_ope.loc[today,:] - result_ope.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_ope.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_operational, outfile=plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".png")

summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
summarypage1 = summarypage1.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
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

data_all_tmp_producttype = []

pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))),color_function=color_function)

result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
result.columns.name = attribute_groupby
result = result.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_all.tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_all.xlsx")

result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
result_total = result_total.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
result_total.columns.name = attribute_groupby
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+"_all.tex",column_format="l|r|rrr|rrr",longtable=True)
result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
result_operational = result_operational.reindex(['DDA-DOMESTIC','IBDDA-DOMESTIC','IBDDA-CAYMAN','IBDDA-EUROPEAN','IBDDA-LONDON','Others','Total'])
result_operational.columns.name = attribute_groupby
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+"_all.tex",column_format="l|r|rrr|rrr",longtable=True)

#### source product type new
#attribute_groupby = column_name_product_typenew
#
#result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
#result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
#result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
#result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
#result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
#result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
#result = result.reindex(['Transact_USD','Transact_EUR','Transact_GBP','Transact_JPY','Transact_CHF','Transact_CAD','Transact_AUD','Transact_OTHER','NIBDDA','IBDDA',"Cayman","Other",'Total'])
#result.columns.name = attribute_groupby
#
#result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
#result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")
#
#result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
#result_total['Total'] = result_total.apply(np.nansum,axis=1)
#result_total = concat_history(column_name_spot,attribute_groupby,result_total)
#summarypage1 = aggregate_summary_table_with_percentiles(result_total)
#summarypage1 = summarypage1.reindex(['Transact USD','Transact EUR','Transact GBP','Transact JPY','Transact CHF','Transact CAD','Transact AUD','Transact OTHER','NIBDDA','IBDDA',"Cayman","Other",'Total'])
#summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
#summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
#summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
#summarypage1diff = summarypage1diff.reindex(['Transact USD','Transact EUR','Transact GBP','Transact JPY','Transact CHF','Transact CAD','Transact AUD','Transact OTHER','NIBDDA','IBDDA',"Cayman","Other",'Total'])
#summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
#summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")
##attrition analysis
#df = result_total.ix[today] - result_total.ix[yesterday]
#df = df[df.index!="Total"]
#df = df[df!=0]
#values = [result_total.loc[yesterday,"Total"]]+df.tolist()
#xtick_names = [formatDate(yesterday)]+df.index.tolist()
#Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_spot, outfile=plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".png")
#
#result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
#result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
#result_ope = concat_history(column_name_operational,attribute_groupby,result_ope)
#summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
#summarypage1 = summarypage1.reindex(['Transact USD','Transact EUR','Transact GBP','Transact JPY','Transact CHF','Transact CAD','Transact AUD','Transact OTHER','NIBDDA','IBDDA',"Cayman","Other",'Total'])
#summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
#summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
#summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
#summarypage1diff = summarypage1diff.reindex(['Transact USD','Transact EUR','Transact GBP','Transact JPY','Transact CHF','Transact CAD','Transact AUD','Transact OTHER','NIBDDA','IBDDA',"Cayman","Other",'Total'])
#summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
#summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")
##attrition analysis
#df = result_ope.ix[today] - result_ope.ix[yesterday]
#df = df[df.index!="Total"]
#df = df[df!=0]
#values = [result_ope.loc[yesterday,"Total"]]+df.tolist()
#xtick_names = [formatDate(yesterday)]+df.index.tolist()
#Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_operational, outfile=plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".png")
#
#summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
#summarypage1 = summarypage1.reindex(['Transact USD','Transact EUR','Transact GBP','Transact JPY','Transact CHF','Transact CAD','Transact AUD','Transact OTHER','NIBDDA','IBDDA',"Cayman","Other",'Total'])
#summarypage1.iloc[:,:7] = summarypage1.iloc[:,:7].applymap(percentages)
#summarypage1.to_latex(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
#summarypage1.to_excel(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".xlsx")
#
#pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=list(set(pivot_attributes).difference(set([attribute_groupby]))),color_function=color_function)
#data_all_tmp_producttype = []

### market segment
attribute_groupby = column_name_market
result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
result = result.reindex(['INVESTMENT/FUND MANAGER','BANK/NON-BANK FINANCIAL','GSE/SOVEREIGN/NONPROFIT','NONFINANCIAL CORPORATION','TRUSTEE','TRADE UNION/PROFESSIONAL ORGANIZATION','UNKNOWN_STT','_Missing','Total'])
result.columns.name = attribute_groupby
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

### Client Size
# process to have client size
attribute_groupby = column_name_parent
sort_by = "MonthAvg"
result_total = aggregate_summary_table(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot,sort_by=sort_by,digit=5)
clients_concentration = result_total.index
clients_concentration = clients_concentration[~clients_concentration.isin(["UnregulatedFunds","_Missing","Total"])]

client_segments = []
client_segments.append(clients_concentration[:10])
client_segments.append(clients_concentration[10:50])
client_segments.append(clients_concentration[50:100])
client_segments.append(clients_concentration[100:200])
client_segments.append(clients_concentration[200:500])
client_segments.append(clients_concentration[500:])
client_segments.append(["UnregulatedFunds"])
client_segments.append(["_Missing"])

labels = ["Clients 1-10", "Clients 11-50", "Clients 51-100", "Clients 101-200", "Clients 201-500", "Clients 501+","UnregulatedFund","_Missing"]
segment = data_all_tmp[column_name_parent].copy()
for i in range(len(client_segments)):
#    i=0
    tmp_names = data_all_tmp[column_name_parent]
    segment[tmp_names.fillna("").isin(client_segments[i])] = labels[i]
#    data_all_tmp[column_name_parent].isin(client_segments[i])
#    segment[data_all_tmp[column_name_parent].isin(client_segments[i])] = labels[i]
segment[pd.isnull(data_all_tmp[column_name_parent])] = "_Missing"
data_all_tmp["Segment"] = segment

attribute_groupby = "Segment"

# spot balance
result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
result.columns.name = attribute_groupby
result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")

result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
result_total = result_total[labels]
result_total['Total'] = result_total.apply(np.nansum,axis=1)
df = result_total.loc[today,:] - result_total.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_total.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_spot, outfile=plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".png")

result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
result_ope = result_ope[labels]
result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
df = result_ope.loc[today,:] - result_ope.loc[yesterday,:]
df = df[df.index!="Total"]
df = df[df!=0]
values = [result_ope.loc[yesterday,"Total"]]+df.tolist()
xtick_names = [formatDate(yesterday)]+df.index.tolist()
Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_operational, outfile=plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".png")

############# Concentration ############
piecolors = ['gold', 'yellowgreen', 'lightcoral', 'lightskyblue', "orange", "palegreen", "mistyrose", "lightslategrey"]
explode = (0.05, 0.05, 0.05, 0,0, 0,0, 0)  # explode 1st slice
# Plot
plt.figure()
pie_label = ["Clients 1-10", "Clients 11-50", "Clients 51-100", "Clients 101-200", "Clients 201-500", "Clients 501+","UnregulatedFunds","Miss Ultimate Parent"]
plt.pie(result_total.loc[today,labels],autopct='%1.1f%%',explode=explode,shadow=True,colors=piecolors,
        labels = pie_label)
plt.axis('equal')
plt.tight_layout()
plt.title("Concentration of Total Deposits",fontsize=30)
pylab.savefig(plot_folder+"/Summary_DepositClientConcentration.png",bbox_inches='tight')
plt.clf()

plt.figure()
tmp = result_ope.loc[today,labels]
plt.pie(tmp[:6],autopct='%1.1f%%',explode=explode[:6],shadow=True,colors=piecolors,
        labels = pie_label[:6])
plt.axis('equal')
plt.tight_layout()
plt.title("Concentration of Operational Deposits",fontsize=30)
pylab.savefig(plot_folder+"/Summary_OperationalDepositClientConcentration.png",bbox_inches='tight')
plt.clf()

############# Concentration - New ############
balancetypes = ["Total_Deposits", "Operational_Deposits"]
for balancetype in balancetypes:
#    balancetype = "Total_Deposits"
    file_name = balancetype + "_By_ParentCompy"
    parent = read_db(file_name)
    total_deposits = parent.apply(np.nansum,axis=0)
    
    top_number = 5
    concentration_5 = parent.loc[clients_concentration[:top_number],:].apply(np.nansum,axis=0)/total_deposits
    top_number = 10
    concentration_10 = parent.loc[clients_concentration[:top_number],:].apply(np.nansum,axis=0)/total_deposits
    top_number = 20
    concentration_20 = parent.loc[clients_concentration[:top_number],:].apply(np.nansum,axis=0)/total_deposits
    top_number = 50
    concentration_50 = parent.loc[clients_concentration[:top_number],:].apply(np.nansum,axis=0)/total_deposits
    top_number = 100
    concentration_100 = parent.loc[clients_concentration[:top_number],:].apply(np.nansum,axis=0)/total_deposits
    top_number = 200
    concentration_200 = parent.loc[clients_concentration[:top_number],:].apply(np.nansum,axis=0)/total_deposits
    top_number = 500
    concentration_500 = parent.loc[clients_concentration[:top_number],:].apply(np.nansum,axis=0)/total_deposits
    
    result_concentration = pd.DataFrame({"Top 5":concentration_5,"Top 10":concentration_10,\
                  "Top 20":concentration_20,"Top 50":concentration_50,\
                  "Top 100":concentration_100,"Top 200":concentration_200,\
                  "Top 500":concentration_500})
    result_concentration = result_concentration[["Top 500","Top 200","Top 100","Top 50","Top 20","Top 10","Top 5"]]
    result_concentration.columns = [x + " ("+str('{:.0%}'.format(round(y,2)))+")" for x,y in zip(result_concentration.columns,result_concentration.loc[max(result_concentration.index),:])]
    
    plt.figure()
    figure = result_concentration.plot(style=styles,linewidth=1.3)
    figure.xaxis.label.set_visible(False)
    figure.yaxis.set_major_formatter(formatter_percentages)
    plt.title("Concentration of " + balancetype.replace("_"," ") + " By Top Clients\n with the Largest Monthly Average Balance\n (Pct in parentheses is the most recent observation)",fontsize=15)
    plt.legend(loc=2, bbox_to_anchor=(1,0.8))
    plt.rcParams.update({'font.size': 12})
    fig = figure.get_figure()
    fig.set_size_inches(8, 6)
    fig.savefig(plot_folder+"/Summary_"+balancetype+"_ClientConcentration.png",bbox_inches='tight')
    plt.clf()


# Statistics for clients
#attribute_groupby = column_name_parent
#balance_type = column_name_spot
#file_name = balance_type+"_By_"+attribute_groupby
#data_total = read_db(file_name).transpose()
#data_total = data_total.reindex_axis(data_total.mean().sort_values(ascending=False).index,axis=1)
#client_std = data_total.rolling(window=22,center=False).std()
#summarypage1 = aggregate_summary_table_with_percentiles(result_total,digit=10)
#summarypage1 = summarypage1.sort_values(by="Average",ascending=False)
#summarypage1['CoeffVar'] = summarypage1['Std']/summarypage1['Average']
#summarypage1['CoeffVar'] = summarypage1['CoeffVar'].replace([np.inf, -np.inf], np.nan).fillna(0)
#summarypage1.to_excel(plot_folder+"/test.xlsx")
#summarypage1[summarypage1['CoeffVar']>0.1]

#increasing
attribute_groupby = column_name_parent
result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,2)
result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
result = result[~result.index.isin(["UnregulatedFund","Total"])]
result = result.head(num_top_funds_to_plot)
result.index = [x.replace("_Missing","Fund Miss Parent Info") for x in result.index]
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
result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
result = result.sort_values("change")
result = result[~result.index.isin(["UnregulatedFund","Total"])]
result = result.head(num_top_funds_to_plot)
result.index = [x.replace("_Missing","Fund Miss Parent Info") for x in result.index]
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
result_total.index = [x.replace("_Missing","Fund Miss Parent Info") for x in result_total.index]
client_largest_average_droping = result_total.index
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="ll|r|rrr|rrr",longtable=True)

result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational,need_aggregate=need_aggregate,top_number=top_number)
result_operational.index = [x.replace("_Missing","Fund Miss Parent Info") for x in result_operational.index]
result_operational = result_operational.join(mapping_code_names,how="left")
null_client = pd.isnull(result_operational["Client"])
result_operational.loc[null_client,"Client"] = result_operational[null_client].index
result_operational.insert(0, "Client",result_operational.pop("Client"))
result_operational.index = list(range(1,len(result_operational)-1))+["",""]
result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="ll|r|rrr|rrr",longtable=True)

##sort by Average Balance
#################### new table ##################
need_aggregate = True
sort_by = "MonthAvg"
##longer table
top_number = 500+2
result_total = aggregate_summary_table(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot,need_aggregate=need_aggregate,top_number=top_number,sort_by=sort_by)
top_clients = result_total.index[~result_total.index.isin(['_Missing','UnregulatedFunds', 'Others', 'Total','Other'])]
result_total_6m = aggregate_summary_table(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot,need_aggregate=False,average_period=282)
result_total = result_total.join(result_total_6m[["MonthAvg","DiffToAvg","%DiffToAvg"]],how="left",rsuffix="6m")

data_orig = read_db("Total_Deposits_By_ParentCompy").transpose()
data_orig_tmp = data_orig[top_clients]
data_orig_tmp["Other"] = data_orig[~data_orig.index.isin(result_total)].apply(np.nansum,1)
percentiles = aggregate_summary_table_with_percentiles(data_orig_tmp.sort_index(ascending=False))["Percentile"][result_total.index]
result_total = pd.concat([result_total,percentiles],axis=1)
result_total.index = [x.replace("_Missing","Fund Miss Parent Info") for x in result_total.index]
client_largest_average = result_total.index
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-3))+["","","",""]
result_total.columns = ['Client', today.strftime("%m-%d"), yesterday.strftime("%m-%d"), 'Change', '%ofChg','MonthAvg', 'DiffToAvg', '%ofChg','6MAvg', 'DiffToAvg', '%ofChg', 'Percentile']
result_total.fillna("").to_excel(plot_folder+"/Summary_Monthly_"+attribute_groupby+".xlsx")

##top 40 table
top_number = num_top_funds_to_plot+2
result_total = aggregate_summary_table(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot,need_aggregate=need_aggregate,top_number=top_number,sort_by=sort_by)
top_clients = result_total.index[~result_total.index.isin(['_Missing','UnregulatedFunds', 'Others', 'Total','Other'])]
result_total_6m = aggregate_summary_table(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot,need_aggregate=False,average_period=282)
result_total = result_total.join(result_total_6m[["MonthAvg","DiffToAvg","%DiffToAvg"]],how="left",rsuffix="6m")

data_orig = read_db("Total_Deposits_By_ParentCompy").transpose()
data_orig_tmp = data_orig[top_clients]
data_orig_tmp["Other"] = data_orig[~data_orig.index.isin(result_total)].apply(np.nansum,1)
percentiles = aggregate_summary_table_with_percentiles(data_orig_tmp.sort_index(ascending=False))["Percentile"][result_total.index]
result_total = pd.concat([result_total,percentiles],axis=1)
result_total.index = [x.replace("_Missing","Fund Miss Parent Info") for x in result_total.index]
client_largest_average = result_total.index
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-3))+["","","",""]
result_total.columns = ['Client', today.strftime("%m-%d"), yesterday.strftime("%m-%d"), 'Change', '%ofChg','MonthAvg', 'DiffToAvg', '%ofChg','6MAvg', 'DiffToAvg', '%ofChg', 'Percentile']
result_total.fillna("").to_latex(plot_folder+"/Summary_Monthly_"+attribute_groupby+".tex",column_format="ll|r|rrr|rrr|rrr|r",longtable=True)

#result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational,need_aggregate=need_aggregate,top_number=top_number)
#result_operational.index = [x.replace("_Missing","Fund Miss Parent Info") for x in result_operational.index]
#result_operational = result_operational.join(mapping_code_names,how="left")
#null_client = pd.isnull(result_operational["Client"])
#result_operational.loc[null_client,"Client"] = result_operational[null_client].index
#result_operational.insert(0, "Client",result_operational.pop("Client"))
#result_operational.index = list(range(1,len(result_operational)-1))+["",""]
#result_operational.columns = ['Client', today.strftime("%m-%d[Total]"), yesterday.strftime("%m-%d[Total]"), 'DiffToBck', '%DiffToBck','MonthAvg', 'DiffToAvg', '%DiffToAvg']
#result_operational.to_latex(plot_folder+"/Summary_Monthly_Operational_"+attribute_groupby+".tex",column_format="ll|r|rrr|rrr",longtable=True)

## pivot table
#pivot_attributes = [column_name_group,column_name_company,column_name_currency,column_name_region,'ProductTypes']
top_clients = top_clients[:num_top_funds_to_plot]
####Excel format
for pivot_attribute in pivot_attributes:
    pivot_attribute = [pivot_attribute]
    result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute,output_excel=False,output_tex=False)
    #select top clients
    result_total = result_total.loc[top_clients,:]
    
    result_total = result_total.join(mapping_code_names,how="left")
    null_client = pd.isnull(result_total["Client"])
    result_total.loc[null_client,"Client"] = result_total[null_client].index
    result_total.insert(0, "Client",result_total.pop("Client"))
    result_total.index = list(range(1,len(result_total)))+[""]
    result_total.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".xlsx")

####LaTex format
#Aggregate smaller clients as Others
data_all_tmp.loc[~data_all_tmp[column_name_parent].isin(top_clients),column_name_parent] = 'Others'

pivot_attribute = [column_name_company]
result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute,output_excel=False,output_tex=False)
result_total = round(result_total/1e9,2)
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
#result_total.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".xlsx")
result_total = result_total.loc[list(top_clients)+["Others","Sum"],:]
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".tex",column_format='l'+'r'*len(result_total.columns),longtable=True)

pivot_attribute = [column_name_group]
result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute,output_excel=False,output_tex=False)
result_total = round(result_total/1e9,2)
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
#result_total.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".xlsx")
result_total = result_total.loc[list(top_clients)+["Others","Sum"],:]
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
if FED_ENVIRONMENT:
    result_total.columns = ['Client', 'Excluded', 'INTRADAY','MDHigh', 'MDLow', 'Non-Disc', 'Sum']
else:
    result_total.columns = ['Client', 'Excluded', 'MissParentInfo', 'INTRADAY','MDHigh', 'MDLow', 'Non-Disc', 'Sum']

result_total.fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".tex",column_format='l'+'r'*len(result_total.columns),longtable=True)

pivot_attribute = [column_name_currency]
result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute,output_excel=False,output_tex=False)
result_total = round(result_total/1e9,2)
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
#result_total.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".xlsx")
result_total = result_total.loc[list(top_clients)+["Others","Sum"],:]
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".tex",column_format='l'+'r'*len(result_total.columns),longtable=True)

#highly customized tables
result_total, pivot_diff = pivot_tables_client(data_all_tmp)
result_total = result_total.loc[top_clients,:]
pivot_diff = pivot_diff.loc[top_clients,:]
client_name = mapping_code_names.loc[top_clients,"Client"]
client_name[client_name.isnull()] = top_clients[client_name.isnull()]
result_total.index = client_name

fig = render_mpl_table(result_total, data_color = pivot_diff.fillna(0), func=color_function_client, col_width=2.0,row_height=0.4)
fig.savefig(plot_folder+"/Summary_DailyValue_Diff_"+attribute_groupby+"_by_"+pivot_attribute[0]+".png",bbox_inches='tight')

pivot_attribute = [column_name_region]
result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute,output_excel=False,output_tex=False)
result_total = round(result_total/1e9,2)
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
#result_total.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".xlsx")
result_total = result_total.loc[list(top_clients)+["Others","Sum"],:]
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".tex",column_format='l'+'r'*len(result_total.columns),longtable=True)

pivot_attribute = ['ProductTypes']
result_total = pivot_tables(data_all_tmp,index=attribute_groupby,pivot_attributes=pivot_attribute,output_excel=False,output_tex=False)
result_total = round(result_total/1e9,2)
result_total = result_total.join(mapping_code_names,how="left")
null_client = pd.isnull(result_total["Client"])
result_total.loc[null_client,"Client"] = result_total[null_client].index
#result_total.fillna("").to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".xlsx")
result_total = result_total.loc[list(top_clients)+["Others","Sum"],:]
result_total.insert(0, "Client",result_total.pop("Client"))
result_total.index = list(range(1,len(result_total)-1))+["",""]
result_total.rename(columns={"DDA-DOMESTIC":'DDA/USA','IBDDA-DOMESTIC':'IBDDA/USA','IBDDA-EUROPEAN':'IBDDA/EU','IBDDA-LONDON':'IBDDA/GBP'},inplace=True)
result_total.fillna("").to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+"_by_"+pivot_attribute[0]+".tex",column_format='l'+'r'*len(result_total.columns),longtable=True)

data_all_tmp = []

if not FED_ENVIRONMENT:
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
    
    result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
    result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
    result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
    result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
    result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
    result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
    result.index = [domicile_diction[x] for x in result.index]
    result.columns.name = "Domicile"
    result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
    result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")
    
    #result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
    #result_total.index = [domicile_diction[x] for x in result_total.index]
    #result_total.columns.name = attribute_groupby
    #result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
    #result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
    #result_operational.index = [domicile_diction[x] for x in result_operational.index]
    #result_operational.columns.name = attribute_groupby
    #result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
    
    data_all_tmp = []
    
    ###Investment Style
    attribute_groupby = column_name_stylegroup
    reorder_index = ["CASH","CASH - MONEY MARKETS","DOMESTIC BOND","GLOBAL BOND","BALANCED","DOMESTIC EQUITY","GLOBAL EQUITY","GROWTH","ALTERNATIVE","MUTUAL FUNDS HOLDING","_Missing","Total"]
    data_all_tmp = data_all[attribute_aggs+[attribute_groupby,"AS_OF_DATE"]]
    
    result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
    result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
    result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
    result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
    result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
    result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
    result = result.reindex(reorder_index)
    result.columns.name = attribute_groupby
    result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
    result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")
    
    result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
    result_total['Total'] = result_total.apply(np.nansum,axis=1)
    result_total = concat_history(column_name_spot,attribute_groupby,result_total)
    summarypage1 = aggregate_summary_table_with_percentiles(result_total)
    summarypage1 = summarypage1.reindex(reorder_index)
    summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
    summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
    summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
    summarypage1diff = summarypage1diff.reindex(reorder_index)
    summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
    summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")
    #attrition analysis
    df = result_total.loc[today,:] - result_total.loc[yesterday,:]
    df = df[df.index!="Total"]
    df = df[df!=0]
    values = [result_total.loc[yesterday,"Total"]]+df.tolist()
    xtick_names = [formatDate(yesterday)]+df.index.tolist()
    Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_spot, outfile=plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".png")
    
    result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
    result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
    result_ope = concat_history(column_name_operational,attribute_groupby,result_ope)
    summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
    summarypage1 = summarypage1.reindex(reorder_index)
    summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
    summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
    summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
    summarypage1diff = summarypage1diff.reindex(reorder_index)
    summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
    summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")
    #attrition analysis
    df = result_ope.loc[today,:] - result_ope.loc[yesterday,:]
    df = df[df.index!="Total"]
    df = df[df!=0]
    values = [result_ope.loc[yesterday,"Total"]]+df.tolist()
    xtick_names = [formatDate(yesterday)]+df.index.tolist()
    Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_operational, outfile=plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".png")
    
    summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
    summarypage1 = summarypage1.reindex(reorder_index)
    summarypage1.iloc[:,:7] = summarypage1.iloc[:,:7].applymap(percentages)
    summarypage1.to_latex(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
    summarypage1.to_excel(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".xlsx")
    
    result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
    result_total.columns.name = attribute_groupby
    result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
    result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
    result_operational.columns.name = attribute_groupby
    result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
    
    data_all_tmp = []
    
    ###Business units
    attribute_groupby = column_name_division
    reorder_index = ['USIS','GS APAC','GS UKEMEA','GS LUXEMBOURG','GS IRELAND','GS GERMANY','GS ITALY','IIS','WMS','AIS','OTHER','_Missing','Total']
    data_all_tmp = data_all[attribute_aggs+[attribute_groupby,"AS_OF_DATE"]]
    
    result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot)
    result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
    result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational)
    result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
    result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
    result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
    result = result.reindex(reorder_index)
    result.columns.name = attribute_groupby
    result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
    result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")
    
    result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
    result_total['Total'] = result_total.apply(np.nansum,axis=1)
    result_total = concat_history(column_name_spot,attribute_groupby,result_total)
    summarypage1 = aggregate_summary_table_with_percentiles(result_total)
    summarypage1 = summarypage1.reindex(reorder_index)
    summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
    summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
    summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
    summarypage1diff = summarypage1diff.reindex(reorder_index)
    summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
    summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")
    #attrition analysis
    df = result_total.loc[today,:] - result_total.loc[yesterday,:]
    df = df[df.index!="Total"]
    df = df[df!=0]
    values = [result_total.loc[yesterday,"Total"]]+df.tolist()
    xtick_names = [formatDate(yesterday)]+df.index.tolist()
    Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_spot, outfile=plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".png")
    
    result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
    result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
    result_ope = concat_history(column_name_operational,attribute_groupby,result_ope)
    summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
    summarypage1 = summarypage1.reindex(reorder_index)
    summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
    summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
    summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
    summarypage1diff = summarypage1diff.reindex(reorder_index)
    summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
    summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")
    #attrition analysis
    df = result_ope.loc[today,:] - result_ope.loc[yesterday,:]
    df = df[df.index!="Total"]
    df = df[df!=0]
    values = [result_ope.loc[yesterday,"Total"]]+df.tolist()
    xtick_names = [formatDate(yesterday)]+df.index.tolist()
    Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_operational, outfile=plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".png")
    
    summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
    summarypage1 = summarypage1.reindex(reorder_index)
    summarypage1.iloc[:,:7] = summarypage1.iloc[:,:7].applymap(percentages)
    summarypage1.to_latex(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
    summarypage1.to_excel(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".xlsx")
    
    result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
    result_total.columns.name = attribute_groupby
    result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
    result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
    result_operational.columns.name = attribute_groupby
    result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

####Business Lines
#attribute_groupby = column_name_bu
#reorder_index = ['USIS','GS Offshore','GS Onshore','Onshore','IIS Summary','IS Canada','Other','Total']
#data_all_tmp = data_all[attribute_aggs+[attribute_groupby,"AS_OF_DATE"]]
#
##read mapping table
#table_name = "Mapping_RESPCenter_BU"
#bu_mapping = read_db(table_name)["BU_LVL_4_DESC"].to_dict()
#data_all_tmp[column_name_bu].replace(bu_mapping,inplace=True)
#group_agg=["Human Resources Lvl4","Enterprise Risk Mgmt Lvl4","SSGA Executive Management","GlobalLink","GL Exchange",7923020,"Glob Operations","SSGA  Investment  Management","Sales & Trading and Research","Wealth Manager Services","Insurance Summ","Investment  Portfolio","Corporate Finance","Global Institutional Group","Portfolio  Solutions","Controllers","Administrative Svc","Securities  Finance","GS Management And Other","Credit Finance","GM COO","Mgmt & Other","GM Management & Other"]
#result_total = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_spot,need_aggregate=True,group_agg=group_agg)
#result_total.columns = [today.strftime("%m-%d")+"[Total]",yesterday.strftime("%m-%d")+"[Total]","change"]
#result_operational = aggregate_summary_table_format_combined(data_all_tmp,today,yesterday,attribute_groupby,column_name_operational,need_aggregate=True,group_agg=group_agg)
#result_operational.columns = [today.strftime("%m-%d")+"[Opertnl]",yesterday.strftime("%m-%d")+"[Opertnl]","change"]
#result = round(result_total.join(result_operational,how='left',rsuffix=" ")/1e9,1)
#result["Capture Rate"] = (result_operational.iloc[:,0]/result_total.iloc[:,0]).map(percentages)
#result = result.reindex(reorder_index)
#result.columns.name = attribute_groupby
#result.to_latex(plot_folder+"/Summary_Intraday_"+attribute_groupby+".tex",column_format="l|rrr|rrr|r",longtable=True)
#result.to_excel(plot_folder+"/Summary_Intraday_"+attribute_groupby+".xlsx")
#
#data_all_tmp = data_all[attribute_aggs+[attribute_groupby,"AS_OF_DATE"]]
#result_total = summarizeTable(data_all_tmp,attribute_groupby,column_name_spot).transpose().sort_index(ascending=False)
#result_total['Total'] = result_total.apply(np.nansum,axis=1)
#result_total = concat_history(column_name_spot,attribute_groupby,result_total)
#
#result_total.ix[:,~result_total.columns.isin(bu_mapping.keys())]
#result_total.ix[:,result_total.columns.isin(bu_mapping.keys())]
#result_total.columns = [bu_mapping[x] if x in bu_mapping else group_agg[0] for x in result_total.columns]
#
#summarypage1 = aggregate_summary_table_with_percentiles(result_total)
#summarypage1 = summarypage1.reindex(reorder_index)
#summarypage1.to_latex(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
#summarypage1.to_excel(plot_folder+"/Summary_Intraday_TotalBal_"+attribute_groupby+".xlsx")
#summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_total)
#summarypage1diff = summarypage1diff.reindex(reorder_index)
#summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
#summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".xlsx")
##attrition analysis
#df = result_total.ix[today] - result_total.ix[yesterday]
#df = df[df.index!="Total"]
#df = df[df!=0]
#values = [result_total.loc[yesterday,"Total"]]+df.tolist()
#xtick_names = [formatDate(yesterday)]+df.index.tolist()
#Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_spot, outfile=plot_folder+"/Summary_Intraday_TotalBalChg_"+attribute_groupby+".png")
#
#result_ope = summarizeTable(data_all_tmp,attribute_groupby,column_name_operational).transpose().sort_index(ascending=False)
#result_ope['Total'] = result_ope.apply(np.nansum,axis=1)
#result_ope = concat_history(column_name_operational,attribute_groupby,result_ope)
#summarypage1 = aggregate_summary_table_with_percentiles(result_ope)
#summarypage1 = summarypage1.reindex(reorder_index)
#summarypage1.to_latex(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
#summarypage1.to_excel(plot_folder+"/Summary_Intraday_OpeBal_"+attribute_groupby+".xlsx")
#summarypage1diff = aggregate_summary_table_with_percentiles_and_changes(result_ope)
#summarypage1diff = summarypage1diff.reindex(reorder_index)
#summarypage1diff.to_latex(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".tex",column_format="lrrrrrr",longtable=True)
#summarypage1diff.to_excel(plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".xlsx")
##attrition analysis
#df = result_ope.ix[today] - result_ope.ix[yesterday]
#df = df[df.index!="Total"]
#df = df[df!=0]
#values = [result_ope.loc[yesterday,"Total"]]+df.tolist()
#xtick_names = [formatDate(yesterday)]+df.index.tolist()
#Waterfall(values, xtick_names, fig_size=(11, 6), xticks_fontsize=15,plot_title="Attrition By "+attribute_groupby,plot_X="",plot_Y=column_name_operational, outfile=plot_folder+"/Summary_Intraday_OpeBalChg_"+attribute_groupby+".png")
#
#summarypage1 = aggregate_summary_table_with_percentiles(result_ope/result_total*1e9,digit=3)
#summarypage1 = summarypage1.reindex(reorder_index)
#summarypage1.iloc[:,:7] = summarypage1.iloc[:,:7].applymap(percentages)
#summarypage1.to_latex(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".tex",column_format="lrrrrrrrr",longtable=True)
#summarypage1.to_excel(plot_folder+"/Summary_Intraday_CaptureRate_"+attribute_groupby+".xlsx")
#
#result_total = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_spot)
#result_total.columns.name = attribute_groupby
#result_total.to_latex(plot_folder+"/Summary_Comparison_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)
#result_operational = aggregate_summary_table(data_all_tmp,currentsnapshot,benchmarkdate,attribute_groupby,column_name_operational)
#result_operational.columns.name = attribute_groupby
#result_operational.to_latex(plot_folder+"/Summary_Comparison_Operational_"+attribute_groupby+".tex",column_format="l|r|rrr|rrr",longtable=True)

data_all_tmp = []
### End of Analysis by Attributes

### FX Impact table in production
#fx_summary = read_db("FX_RATES_Summary")
#balance_by_currency = read_db("BALANCEBYCURRENCY_Summary")
#fx_summary = fx_summary.loc[balance_by_currency.index,:]
#
#if today in balance_by_currency.columns:
#    balance_today = balance_by_currency[today]
#    
#    date_yesterday = max(balance_by_currency.columns[balance_by_currency.columns < today])
#    date_1w = max(balance_by_currency.columns[balance_by_currency.columns <= today-datetime.timedelta(days=7)])
#    date_1m = max(balance_by_currency.columns[balance_by_currency.columns <= today-datetime.timedelta(days=365/12*1)])
#    date_3m = max(balance_by_currency.columns[balance_by_currency.columns <= today-datetime.timedelta(days=365/12*3)])
#    date_6m = max(balance_by_currency.columns[balance_by_currency.columns <= today-datetime.timedelta(days=365/12*6)])
#    date_1y = max(balance_by_currency.columns[balance_by_currency.columns <= today-datetime.timedelta(days=365)])
#    
#    fx_impact = fx_summary[[date_yesterday,date_1w,date_1m,date_3m,date_6m,date_1y]].multiply(balance_today,axis="index").subtract(fx_summary[today]*balance_today,axis="index")
#    
#    major_impact = fx_impact[fx_impact.index.isin(["AUD","CAD","CHF","EUR","GBP","JPY"])]
#    other_impact = pd.DataFrame({"Others":fx_impact[~fx_impact.index.isin(["AUD","CAD","CHF","EUR","GBP","JPY"])].apply(np.nansum,axis=0)}).transpose()
#    tot_impact = pd.DataFrame({"Total":fx_impact.apply(np.nansum,axis=0)}).transpose()
#    
#    fx_impact = round(major_impact.append(other_impact).append(tot_impact)/1e9,2)
#    
#    a=[x.strftime("%m%d%y") for x in fx_impact.columns]
#    b=['1D','1W','1M','3M','6M','1Y']
#    fx_impact.columns = [x+"-"+y for x,y in zip(a,b)]
#    fx_impact.index.name = None
#    fx_impact.to_latex(plot_folder+"/FX_Impact.tex",column_format='lrrrrrr',longtable=True)
#    
#    # balance and fx rate info
#    balance_local = round(balance_by_currency.loc[["AUD","CAD","CHF","EUR","GBP","JPY"],[today]]/1e9,2)
#    balance_local.index.name=None
#    balance_local.columns=[x.strftime("%m%d%y")+"-Balance" for x in balance_local.columns]
#    
#    fx_rates = fx_summary.loc[["AUD","CAD","CHF","EUR","GBP","JPY"],[today,date_yesterday,date_1w,date_1m,date_3m,date_6m,date_1y]]
#    a=[x.strftime("%m%d%y") for x in fx_rates.columns]
#    b=['0D','1D','1W','1M','3M','6M','1Y']
#    fx_rates.columns = [x+"-"+y for x,y in zip(a,b)]
#    fx_rates.index.name = None
#    fx_info = balance_local.join(fx_rates)
#    fx_info.to_latex(plot_folder+"/FX_Info.tex",column_format='lr|rrrrrrr',longtable=True)
#else:
#    print("Today's balance is not updated. FX impact is not assessed.")
#
### new dev for FX
#if TEST_ENVIRONMENT:
#    balance_by_currency_usd = read_db("BALANCEBYCURRENCY_USD_Summary")
#    major_currencies = ["USD","AUD","CAD","CHF","EUR","GBP","JPY","Others"]
#    fx_start_date = datetime.datetime.strptime('2018-04-01 00:00:00', '%Y-%m-%d %H:%M:%S')
#    fx_end_date = datetime.datetime.strptime('2018-04-30 00:00:00', '%Y-%m-%d %H:%M:%S')
#    
#    pct_change = fx_summary.loc[fx_summary.index.isin(major_currencies),(fx_summary.columns>=fx_start_date)&(fx_summary.columns<=fx_end_date)].transpose().pct_change()
#    pct_change = pct_change.fillna(0)
#    
#    plt.figure()
#    figure = pct_change.dropna().plot(style=styles,linewidth=1.3)
#    figure.xaxis.label.set_visible(False)
#    plt.title("Daily Change of Currency Value Relative To USD",fontsize=15)
#    plt.legend(loc=2, bbox_to_anchor=(1,0.8))
#    plt.rcParams.update({'font.size': 12})
#    fig = figure.get_figure()
#    fig.set_size_inches(8, 6)
#    fig.savefig(plot_folder+"/FX_Impact_Daily_FX_Change.png",bbox_inches='tight')
#    plt.clf()
#    pct_change.to_excel(plot_folder+"/FX_Impact_Daily_FX_Change.xlsx")
#    
#    cum_change = fx_summary.loc[fx_summary.index.isin(major_currencies),(fx_summary.columns>=fx_start_date)&(fx_summary.columns<=fx_end_date)].transpose()
#    cum_change = cum_change/cum_change.loc[min(cum_change.index),:]
#    plt.figure()
#    figure = cum_change.dropna().plot(style=styles,linewidth=1.3)
#    figure.xaxis.label.set_visible(False)
#    plt.title("Cumulative Change of Currency Value Relative To USD",fontsize=15)
#    plt.legend(loc=2, bbox_to_anchor=(1,0.8))
#    plt.rcParams.update({'font.size': 12})
#    fig = figure.get_figure()
#    fig.set_size_inches(8, 6)
#    fig.savefig(plot_folder+"/FX_Impact_Cumulative_FX_Change.png",bbox_inches='tight')
#    plt.clf()
#    cum_change.to_excel(plot_folder+"/FX_Impact_Cumulative_FX_Change.xlsx")
#
#    balance = balance_by_currency_usd.loc[balance_by_currency_usd.index.isin(major_currencies[1:]),(balance_by_currency_usd.columns>=fx_start_date)&(balance_by_currency_usd.columns<=fx_end_date)].transpose()
#    plt.figure()
#    figure = balance.dropna().plot(style=styles,linewidth=1.3)
#    figure.xaxis.label.set_visible(False)
#    plt.title("Balance by Currency in US Dollar equivalent",fontsize=15)
#    plt.legend(loc=2, bbox_to_anchor=(1,0.8))
#    plt.rcParams.update({'font.size': 12})
#    fig = figure.get_figure()
#    fig.set_size_inches(8, 6)
#    fig.savefig(plot_folder+"/FX_Impact_Daily_FX_Balance.png",bbox_inches='tight')
#    plt.clf()
#    balance.to_excel(plot_folder+"/FX_Impact_Daily_FX_Balance.xlsx")
#    
#    # assess the impact of FX currency
#    balance_with_updated_rates =(balance_by_currency.loc[:,(balance_by_currency.columns>=fx_start_date)&(balance_by_currency.columns<=fx_end_date)]*
#     fx_summary.loc[:,(fx_summary.columns>=fx_start_date)&(fx_summary.columns<=fx_end_date)]).apply(np.nansum,0)
#
#    balance = balance_by_currency.loc[:,(balance_by_currency.columns>=fx_start_date)&(balance_by_currency.columns<=fx_end_date)]
#    rates = fx_summary[max(balance.columns)]
#    for column in balance:
#        balance[column] = balance[column]*rates
#    diff_pct = ((balance_with_updated_rates-balance.apply(np.nansum,0))/balance_with_updated_rates)
#    diff_abs = (balance_with_updated_rates-balance.apply(np.nansum,0))
#    pd.DataFrame({"Percentage":diff_pct,"Amount":diff_abs}).to_excel(plot_folder+"/FX_Impact_Daily_FX_Impacy.xlsx")
#    pd.DataFrame({"Deposits Before Adj":balance_with_updated_rates,"Deposits with Constant FX rate":balance.apply(np.nansum,0)}).to_excel(plot_folder+"/FX_Impact_Daily_Deposit.xlsx")
#
#
##    fig, axs = plt.subplots(ncols=2, sharey=True)
##    diff_pct.plot(ax=axs[1],style=styles2,linewidth=1.3,secondary_y=True)
##    diff_abs.plot(ax=axs[0],style=styles,linewidth=1.3)
##    ax = axs[0]
##    ax.set_title("Impact of FX Rates in USD")
##    ax.yaxis.set_major_formatter(formatter_billions)
##    ax = axs[1]
##    ax.set_title("Impact of FX Rates in Percentage")
##    ax.tick_params(labelsize=20) 
##    ax.xaxis.label.set_visible(False)
##    fig.set_size_inches(8,6)
##    fig.savefig(plot_folder+"/FX_Impact_Daily_FX_Impacy.png",bbox_inches='tight')
##    plt.clf()
#
#    balance = balance_by_currency.loc[:,(balance_by_currency.columns>=fx_start_date)&(balance_by_currency.columns<=fx_end_date)]
#    for column in balance:
#        last_day = max(balance_by_currency.columns[balance_by_currency.columns<column])
#        balance[column] = balance[column]*fx_summary[last_day]
#    balance_prior_rate = balance.apply(sum,0)
#    
#    pd.DataFrame({"Balance With Current Rate":balance_with_updated_rates,"Balance With Prior Day Rate":balance.apply(sum,0)}).to_excel(plot_folder+"/FX_Impact_Daily_Assessment.xlsx")
##end of new dev
#fx_summary=[]
#balance_by_currency=[]
#fx_impact=[]

def previous_quarter(ref):
    if ref.month < 4:
        return datetime.datetime.strptime(str(ref.year-1)+"-12-31 00:00:00", "%Y-%m-%d %H:%M:%S")
    elif ref.month < 7:
        return datetime.datetime.strptime(str(ref.year)+"-3-31 00:00:00", "%Y-%m-%d %H:%M:%S")
    elif ref.month < 10:
        return datetime.datetime.strptime(str(ref.year)+"-6-30 00:00:00", "%Y-%m-%d %H:%M:%S")
    return datetime.datetime.strptime(str(ref.year)+"-9-30 00:00:00", "%Y-%m-%d %H:%M:%S")

###plots of history
def cumulative_change_table(data):
#    data = data_total
    date_yesterday = max(data.index[data.index < today])
    date_1w = max(data.index[data.index <= today-datetime.timedelta(days=7)])
    date_1m = max(data.index[data.index <= today-datetime.timedelta(days=365/12*1)])
    date_3m = max(data.index[data.index <= today-datetime.timedelta(days=365/12*3)])
    date_6m = max(data.index[data.index <= today-datetime.timedelta(days=365/12*6)])
    date_1y = max(data.index[data.index <= today-datetime.timedelta(days=365)])
    first = today.replace(day=1)
    date_last_monthend = first - datetime.timedelta(days=1)
    while date_last_monthend not in data.index:
        date_last_monthend = date_last_monthend - datetime.timedelta(days=1)
    
    date_last_qtrend = max(data.index[data.index <= previous_quarter(today)])
    
    cum_change = -1*data.loc[[date_yesterday,date_1w,date_1m,date_3m,date_6m,date_1y,date_last_monthend,date_last_qtrend],:].sub(data.loc[today,:],axis=1)
    cum_change = cum_change.transpose()

    a=[x.strftime("%m%d%y") for x in cum_change.columns]
    b=["1D","1W","1M","3M","6M","1Y","LME","LQE"]
    cum_change.columns = [x+"-"+y for x,y in zip(a,b)]
        
    cum_change = cum_change.sort_values(date_last_qtrend.strftime("%m%d%y")+"-LQE")
    cum_change.loc['Total',:] = cum_change.apply(np.nansum,0)
    
    cum_change = round(cum_change/1e9,1)
    return(cum_change)
    
for attribute in [column_name_group,column_name_company,column_name_currency,column_name_region,'ProductTypes']:
#for attribute in [column_name_currency]:
#    attribute = column_name_currency
    balance_type = column_name_spot
    file_name = balance_type+"_By_"+attribute
    data_total = read_db(file_name).transpose()
    if FED_ENVIRONMENT:
        if 'Funds Miss Model Info' in data_total.columns:
            data_total['Excluded Funds'] = data_total['Excluded Funds'] + data_total['Funds Miss Model Info']
            del data_total['Funds Miss Model Info']
        if 'Funds Miss Region Info' in data_total.columns:
            data_total['Excluded Funds'] = data_total['Excluded Funds'] + data_total['Funds Miss Region Info']
            del data_total['Funds Miss Region Info']

    data_total = data_total.reindex_axis(data_total.mean().sort_values(ascending=False).index,axis=1)
    
    data_plot = data_total.copy()
    
    data_plot.columns=[a_+"("+b_+")" for a_, b_ in zip(data_plot.columns, data_plot.div(data_plot.sum(axis=1),axis=0)[-252:].mean().map(percentages))]
    plt.figure()
    figure = data_plot.plot(style=styles,linewidth=1.3)
    figure.yaxis.set_major_formatter(formatter_billionsinteger)
    figure.xaxis.label.set_visible(False)
    plt.title("Trend of Total Deposit by "+attribute+" in USD (Average Percentage in Parentheses)",fontsize=15)
    plt.legend(loc=2, bbox_to_anchor=(1,0.8))
    plt.rcParams.update({'font.size': 12})
    fig = figure.get_figure()
    fig.set_size_inches(10, 8)
    fig.savefig(plot_folder+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+".png",bbox_inches='tight')
    data_total.to_excel(plot_folder+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+".xlsx")
    if attribute == column_name_currency:
        fig.savefig(weekly_table_dir+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+"_AsOf_"+today.strftime("%Y%m%d")+".png",bbox_inches='tight')
        data_total.to_excel(weekly_table_dir+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+".xlsx")
    plt.clf()

    cum_change = cumulative_change_table(data_total)
    cum_change.to_latex(plot_folder+"/History_CumulativeChange_By_Attribute_"+balance_type+"_"+attribute+".tex",column_format="lrrrrrrrr",longtable=True)
    
    balance_type = column_name_operational
    file_name = balance_type+"_By_"+attribute
    data_operational = read_db(file_name).transpose()
    data_operational = data_operational.reindex_axis(data_total.columns,axis=1)
    
    data_plot = data_operational.copy()
    data_plot.columns=[a_+"("+b_+")" for a_, b_ in zip(data_plot.columns, data_plot.div(data_plot.sum(axis=1),axis=0)[-252:].mean().map(percentages))]
    plt.figure()
    figure = data_plot.plot(style=styles,linewidth=1.3)
    figure.yaxis.set_major_formatter(formatter_billionsinteger)
    figure.xaxis.label.set_visible(False)
    plt.title("Trend of Operational Deposit by "+attribute+" in USD (Average Percentage in Parentheses)",fontsize=15)
    plt.legend(loc=2, bbox_to_anchor=(1,0.8))
    plt.rcParams.update({'font.size': 12})
    fig = figure.get_figure()
    fig.set_size_inches(10, 8)
    fig.savefig(plot_folder+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+".png",bbox_inches='tight')
    data_operational.to_excel(plot_folder+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+".xlsx")
    if attribute == column_name_currency:
        fig.savefig(weekly_table_dir+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+"_AsOf_"+today.strftime("%Y%m%d")+".png",bbox_inches='tight')
        data_operational.to_excel(weekly_table_dir+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+".xlsx")
    plt.clf()
    cum_change = cumulative_change_table(data_operational)
    cum_change.to_latex(plot_folder+"/History_CumulativeChange_By_Attribute_"+balance_type+"_"+attribute+".tex",column_format="lrrrrrrrr",longtable=True)

    balance_type = column_name_excess
    data_excelss = data_total-data_operational
    data_excelss = data_excelss.reindex_axis(data_total.columns,axis=1)

    data_plot = data_excelss.copy()
    data_plot.columns=[a_+"("+b_+")" for a_, b_ in zip(data_plot.columns, data_plot.div(data_plot.sum(axis=1),axis=0)[-252:].mean().map(percentages))]
    plt.figure()
    
    figure = data_plot.plot(style=styles,linewidth=1.3)
    figure.yaxis.set_major_formatter(formatter_billionsinteger)
    figure.xaxis.label.set_visible(False)
    plt.title("Trend of Excess Deposit by "+attribute+" in USD (Average Percentage in Parentheses)",fontsize=15)
    plt.legend(loc=2, bbox_to_anchor=(1,0.8))
    plt.rcParams.update({'font.size': 12})
    fig = figure.get_figure()
    fig.set_size_inches(10, 8)
    fig.savefig(plot_folder+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+".png",bbox_inches='tight')
    data_excelss.to_excel(plot_folder+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+".xlsx")
    if attribute == column_name_currency:
        fig.savefig(weekly_table_dir+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+"_AsOf_"+today.strftime("%Y%m%d")+".png",bbox_inches='tight')
        data_excelss.to_excel(weekly_table_dir+"/History_Chart_By_Attribute_"+balance_type+"_"+attribute+".xlsx")
    plt.clf()
    cum_change = cumulative_change_table(data_excelss)
    cum_change.to_latex(plot_folder+"/History_CumulativeChange_By_Attribute_"+balance_type+"_"+attribute+".tex",column_format="lrrrrrrrr",longtable=True)

#########################################################
aggregate_groupby = column_name_parent
result = summarizeTable(data_all,aggregate_groupby,column_name_spot)
result_operational = summarizeTable(data_all,aggregate_groupby,column_name_operational)

result_byid = summarizeTable(data_all,"ID",column_name_spot)
result_operational_byid = summarizeTable(data_all,"ID",column_name_operational)

#get attributes
mapping_number_behavior = data_all[["AS_OF_DATE","ID",column_name_parent,column_name_parentname,column_name_region,column_name_group,'ND_IND']].copy()
mapping_number_behavior[column_name_parentname] = [x.replace("(PARENT)","").replace("CONFIDENTIAL CLIENT","Conf. Client")[:34] for x in mapping_number_behavior[column_name_parentname].fillna("_Missing")]

mapping_fund_behavior = data_all[["AS_OF_DATE","ID","FUND_NUM","FUND_NAME",column_name_parent,column_name_parentname,column_name_region,column_name_group,'ND_IND']].copy()
mapping_fund_behavior[column_name_parentname] = [x.replace("(PARENT)","").replace("CONFIDENTIAL CLIENT","Conf. Client")[:34] for x in mapping_fund_behavior[column_name_parentname].fillna("_Missing")]

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

date_mapping = pd.ExcelFile(date_mapping_file).parse("Sheet1")
date_mapping.set_index("table_name",inplace=True,drop=True)
date_mapping = date_mapping.dropna()

table_name = "OPSDEP_MASTERATTR"                    #OPSDEP_MASTERATTR table - style
mapping_attribute = read_db(table_name,formatDate(date_mapping.loc[today,table_name]))
mapping_attribute = mapping_attribute[["FUND_ID","FUND_NAME","DIVISION_GROUP","MARKET_SEGMENT_DESC","STYLE_GROUP"]].drop_duplicates()

#payment transaction info
payment_parent = read_db("Payment_By_ParentRegion")
balance_tabl = read_db("Balance_By_ParentRegion")
###Trend of top funds selected
def transform_string_to_legend(string,code):
    return(string.replace(code,code+"/")[:-1]+"/"+string.replace(code,code+"/")[-1])
def extend_ts(spot,total_deposit_by_parent):
    spot_extended = total_deposit_by_parent.loc[spot.columns,:].transpose()
    flag_is_identical = all(abs(spot_extended.loc[spot.index,:]-spot.transpose()) < 1)
    if flag_is_identical:
        return(spot_extended[spot_extended.index <= max(spot.index)])
    else:
        return(spot)

def graph_parent_list(fund_list,plot_threshold=True):
    fund_list = fund_list[~pd.isnull(fund_list)]
    fund_list = fund_list[~fund_list.isin(['Fund Miss Parent Info','UnregulatedFunds','Other'])]
    
    for i in range(len(fund_list)):
#        i = 0
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

        constraint = binding_constraint_by_parent.loc[fund_list[i],:][binding_constraint_by_parent.columns <= max(spot.index)]
        
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
        
#        fig, ax = plt.subplots()
#        client_std[fund_list[i]].plot(ax=ax,style=styles)
#        ax.yaxis.set_major_formatter(formatter_billions)
#        ax.tick_params(labelsize=20)
#        ax.xaxis.label.set_visible(False)
#        ax.set_title("Std of Total Variance",fontsize=15)
#        fig.suptitle("Trend of Deposit Balance of " + fund_list[i] + " (in USD)", fontsize=20)
#        fig.set_size_inches(12,9)
        
        subs = mapping_number_behavior[mapping_number_behavior[column_name_parent]==fund_list[i]]
        subs = subs.dropna()
#        subs_unique = subs.drop_duplicates(["ID",column_name_group ,column_name_parentname, column_name_parent])
        subs_unique = subs.drop_duplicates(["ID" ,column_name_parentname, column_name_parent],keep='first')
        subs_unique = subs_unique.sort_values(["ID"])
        del subs_unique['ID']

        subs_payments = payment_parent.loc[payment_parent.index.isin(subs_unique.index),:]
        subs_balance = balance_tabl.loc[balance_tabl.index.isin(subs_unique.index),:]

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
                client_name = mapping_code_names.loc[fund_list[i],'Client']
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

            balances = subs_balance.loc[funds].transpose()
            spot = result_byid.loc[funds].transpose()
            spot = extend_ts(spot,total_deposit_by_parent_region)
            #extend spot 
            if set(balances.columns)==set(spot.columns):
                spot = balances[balances.index<min(spot.index)].append(spot)
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
                
                behaviors = subs[["AS_OF_DATE","ID","BehaviorGroup"]].drop_duplicates().pivot(index="AS_OF_DATE",columns="ID",values="BehaviorGroup")
                behaviors = extend_ts(behaviors,behavior_group_by_parent_region)
                
                fig, ax = plt.subplots()
                spot.plot(ax=ax, style=styles)
                ope.plot(ax=ax, style=styles2)
                constraints.groupby("ID").plot(kind = "area",color="green",ax=ax,alpha=0.1,label=None,legend=False)
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
                
            #balance to payment figure
            payments = subs_payments.loc[funds].transpose()
            fig, ax = plt.subplots()
            payments.plot(ax=ax, style=styles)
            ax.legend(loc='best')
            ax.yaxis.set_major_formatter(formatter_billions)
            ax.tick_params(labelsize=20)
            ax.xaxis.label.set_visible(False)
            plt.title("Outflow Amount of " + fund_list[i])
            fig.set_size_inches(12,9)
            fig.savefig(plot_folder+"/Trend_Payment_"+fund_list[i]+".png",bbox_inches='tight')
            plt.clf()

            fig, ax = plt.subplots()
            (payments.fillna(0)/spot).dropna(axis=0).rolling(21*3).apply(np.nanmean).dropna().plot(ax=ax, style=styles)
            ax.legend(loc='best')
            ax.tick_params(labelsize=20)
            ax.xaxis.label.set_visible(False)
            plt.title("Payment-to-Balance Ratio of " + fund_list[i] + " (Rolling 3M Average)")
            fig.set_size_inches(12,9)
            fig.savefig(plot_folder+"/Trend_PTB_Ratio_"+fund_list[i]+".png",bbox_inches='tight')
            plt.clf()

def graph_fund_list(fund_list):
    balance_type = column_name_spot
    file_name = balance_type+"_By_"+"FUND_NUM"
    data_total = read_db(file_name).transpose()

    balance_type = column_name_operational
    file_name = balance_type+"_By_"+"FUND_NUM"
    data_operational = read_db(file_name).transpose()

    fund_list = fund_list[~pd.isnull(fund_list)]
    fund_list = fund_list[~fund_list.isin(['Fund Miss Parent Info','UnregulatedFunds','Other','Others'])]
    
    for i in range(len(fund_list)):
#        i = 6
        subs = mapping_fund_behavior.loc[mapping_fund_behavior[column_name_parent]==fund_list[i],["FUND_NUM","FUND_NAME"]].drop_duplicates()

        if len(subs) > 0:
            funds = subs["FUND_NUM"].unique()
            funds = funds[~pd.isnull(funds)]
                          
            # analysis of entire clients' deposit base
            spot = data_total[funds]

            subtotal = spot.apply(np.nansum,axis=1)
            subtotal.name = "Total"
            total = pd.concat([spot,subtotal],axis=1).sort_index(ascending=False)
            summarypage_spot = aggregate_summary_table_with_percentiles(total,attributes=total.columns,digit=4).sort_values("change")
            summarypage_spot.join(mapping_attribute[mapping_attribute["FUND_ID"].isin(funds)].set_index("FUND_ID"),how="left").to_excel(plot_folder+"/Trend_"+aggregate_groupby+"_Total_"+fund_list[i]+"_funddetails.xlsx")
            
            operational = data_operational[funds]
            
            # get funds that represent 5% or more of total client
            share = spot.apply(np.nansum,axis=0)/np.nansum(spot)
            share = share.sort_values(ascending=False)
            funds = share[share>0.01][:5].index

            mapping_attribute[mapping_attribute["FUND_ID"].isin(funds)].to_latex(plot_folder+"/Trend_"+aggregate_groupby+"_"+fund_list[i]+"_funddetails.tex",index=False,longtable=True)

            if len(funds) > 0:                
                fig, ax = plt.subplots()
                spot[funds].plot(ax=ax, style=styles)
                ax.yaxis.set_major_formatter(formatter_millions)
                ax.tick_params(labelsize=20)
                ax.xaxis.label.set_visible(False)
                plt.title("Total Deposits for the Major Funds of " + fund_list[i] + "in USD Millions")
                fig.set_size_inches(12,9)
                fig.savefig(plot_folder+"/Trend_"+aggregate_groupby+"_"+fund_list[i]+"_funddetails_total.png",bbox_inches='tight')
                plt.clf()

                fig, ax = plt.subplots()
                operational[funds].plot(ax=ax, style=styles)
                ax.yaxis.set_major_formatter(formatter_millions)
                ax.tick_params(labelsize=20)
                ax.xaxis.label.set_visible(False)
                plt.title("Operational Deposit for the Major Funds of " + fund_list[i] + "in USD Millions")
                fig.set_size_inches(12,9)
                fig.savefig(plot_folder+"/Trend_"+aggregate_groupby+"_"+fund_list[i]+"_funddetails_operational.png",bbox_inches='tight')
                plt.clf()

#fund_list = client_largest_average_droping[:int(num_top_funds_to_plot)]
#graph_fund_list(fund_list, plot_threshold = False)
if FED_ENVIRONMENT:
    fund_list = top_clients[:2]
else:
    fund_list = top_clients[:num_top_funds_to_plot]
#fund_list = top_clients[:2]
graph_parent_list(fund_list, plot_threshold = False)
graph_fund_list(fund_list)

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
"\\newcommand\\AsOfDate{"+today.strftime("%Y%m%d") +"}\n"+\
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
"\\newcommand\\TopParentMonthlyAverageBalance{"+",".join(top_clients[:num_top_funds_to_plot])+"}\n"+\
"\\newcommand\\TopParentMonthlyWithdraw{"+",".join(client_largest_average_droping[:num_top_funds_to_plot])+"}\n"+\
"\\newcommand\\TopParent{"+column_name_parent+"}\n"
file.write(command)
file.close()

os.chdir(pdf_dir)
#subprocess.call(['pdflatex', '-output-directory', "./", '-jobname', 'DepositAnalysis_'+day, './document.tex'])

#subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis_'+day, './document.tex'])
#subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis_'+day, './document.tex'])
#if PRODUCTION_ENVRIONMENT:
#    subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis_'+day, './document.tex'])
#    subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis_'+day, './document.tex'])
    
#subprocess.call(['pdflatex', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_'+day, './document_workingcopy.tex'])

subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_AsOf_'+today.strftime("%Y%m%d"), './document_workingcopy.tex'])
subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_AsOf_'+today.strftime("%Y%m%d"), './document_workingcopy.tex'])
subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_AsOf_'+today.strftime("%Y%m%d"), './document_workingcopy.tex'])
subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_AsOf_'+today.strftime("%Y%m%d"), './document_workingcopy.tex'])

##adhoc analysis
#subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_adhoc_AsOf_'+today.strftime("%Y%m%d"), './document_workingcopy_adhoc.tex'])
#subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_adhoc_AsOf_'+today.strftime("%Y%m%d"), './document_workingcopy_adhoc.tex'])
#subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_adhoc_AsOf_'+today.strftime("%Y%m%d"), './document_workingcopy_adhoc.tex'])
#subprocess.call(['pdflatex', '-interaction', 'nonstopmode', '-output-directory', "./", '-jobname', 'DepositAnalysis(workingcopy)_adhoc_AsOf_'+today.strftime("%Y%m%d"), './document_workingcopy_adhoc.tex'])

end_time = time.time()
print("Execution runtime is "+str(round(end_time-start_time,0))+" seconds")
#Send message to email
import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
#messages = outlook.GetDefaultFolder(5).items
messages = outlook.GetDefaultFolder(5).items

for message in messages:
    if message.Subject == "Daily Deposit Report":
        print(message.Subject)
#        break
        message.Reply()
        message.To = "jchuang@statestreet.com"
        message.Body = "Automatically sent email on "+datetime.datetime.fromtimestamp(end_time).strftime('%Y-%m-%d-%H-%M-%S')+" \n"+\
                        "Today's deposit report is attached.\n"+\
                        "Today total deposit is "+ billions(Reporting_Total_Deposit)+"; operational deposit is "+billions(Reporting_Ope_Deposit)+" \n"+\
                        "Prior total deposit is "+ billions(Reporting_Total_Deposit_Prior_Day)+"; operational deposit is "+billions(Reporting_Ope_Deposit_Prior_Day)+"\n"

        message.Attachments.Remove(2)
        message.Attachments.Remove(1)
        message.Attachments.Remove(0)
        attachment1 = pdf_dir+'/DepositAnalysis(workingcopy)_AsOf_'+today.strftime("%Y%m%d")+'.pdf'
#        attachment1 = pdf_dir+'/DepositAnalysis(workingcopy)_AsOf_'+"20180330"+'.pdf'
        message.Attachments.Add(attachment1)
        message.Send()
