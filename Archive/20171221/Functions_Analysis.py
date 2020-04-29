# -*- coding: utf-8 -*-
"""
Created on Mon Oct 30 09:14:23 2017

@author: e620927
"""

import pandas as pd
import numpy as np

def listisin(first, second):
        second = set(second)
        return [item for item in first if item in second]

def formatDate(date):
    return(date.strftime("%Y%m%d"))

def summarizeTableByAttribute(data_all,attribute_agg,func=np.nansum):
#    attribute_agg = "TOTAL_PRIN_BAL_USD"
#   data_all=balance
#    attribute_agg = ["DAILY_OPERATIONAL_BALANCE_FINAL","DAILY_SPOT_BALANCE_FINAL"]
    if type(attribute_agg) is str:
        data_all_date = data_all.groupby("AS_OF_DATE")
        result = data_all_date.agg({attribute_agg:func})
        return(result)
    elif type(attribute_agg) is list:
        result = pd.DataFrame()
        for att in attribute_agg:
#            att = attribute_agg[0]
            result = pd.concat([result,summarizeTableByAttribute(data_all,att,func)],axis=1)
        return(result)

def summarizeTable(data_all,attribute_groupby,attribute_agg,fillna=True,func=np.nansum):
#    attribute_groupby = "BEHAVIORAL_GROUP"
#    attribute_agg = "id"
    dates = data_all["AS_OF_DATE"].unique()
    dates = dates[dates.sort()][0]
    
    for i in range(len(dates)):
        data_all_date = data_all[data_all["AS_OF_DATE"] == dates[i]]
        if fillna:
            data_all_date[attribute_groupby].fillna("_Missing",inplace=True)
        data_grouped = data_all_date.groupby(attribute_groupby)

        if i == 0:
            result = data_grouped.agg({attribute_agg:func})
        else:
            result = pd.concat([result,data_grouped.agg({attribute_agg:func})],axis=1)

    result.columns = dates
    return(result)

def summarizeTableRatioOfTwoFactor(data_all,attribute_groupby,attribute_agg_nominator,attribute_agg_denominator):
#    attribute_groupby = "BEHAVIORAL_GROUP"
#    attribute_agg_nominator = "DAILY_OPERATIONAL_BALANCE_FINAL"
#    attribute_agg_denominator = "DAILY_SPOT_BALANCE_FINAL"
    dates = data_all["AS_OF_DATE"].unique()
    dates = dates[dates.sort()][0]
    for i in range(len(dates)):
        data_all_date = data_all[data_all["AS_OF_DATE"] == dates[i]]
        
        data_grouped = data_all_date.groupby(attribute_groupby)
        nominator = data_grouped.agg({attribute_agg_nominator:np.nansum})
        denominator = data_grouped.agg({attribute_agg_denominator:np.nansum})
        if i == 0:
            result = nominator[attribute_agg_nominator]/denominator[attribute_agg_denominator]
        else:
            result = pd.concat([result,nominator[attribute_agg_nominator]/denominator[attribute_agg_denominator]],axis=1)

    result.columns = dates
    return(result)
