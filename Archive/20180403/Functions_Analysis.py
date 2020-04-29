# -*- coding: utf-8 -*-
"""
Created on Mon Oct 30 09:14:23 2017

@author: e620927
"""

import pandas as pd
import numpy as np
import datetime

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
#   data_all = data_date
#    attribute_groupby = attribute_groupby
#    attribute_agg = attribute_agg
    dates = data_all["AS_OF_DATE"].unique()
    dates = dates[dates.sort()][0]
    
    for i in range(len(dates)):
        data_all_date = data_all[data_all["AS_OF_DATE"] == dates[i]]
        if fillna:
            data_all_date[attribute_groupby].fillna("_Missing",inplace=True)
#            data_all_date[attribute_groupby] = data_all_date[attribute_groupby].fillna("_Missing")

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

def process_dda_GTRM(data_all):
#    data_all = table_tmp
    data_all.index=data_all['ULT_PARENT_CD']+data_all['REGION']+data_all['ND_IND']
    data_all['AS_OF_DATE'] = [datetime.datetime.strptime(x, '%Y-%m-%d %H:%M:%S') for x in data_all['AS_OF_DATE']]
    
    data_all['ID'] = data_all.index
    
    #rename columns
    column_name_operational_orig = "DAILY_OPERATIONAL_BAL_USD"
    column_name_spot_orig = "DAILY_SPOT_BAL_USD"
    column_name_excess_orig = "DAILY_EXCESS_BAL_USD"
    
    column_name_company_orig = "IFS_COMPANY"
    column_name_bu_orig = "PROGRESSION_RESP_CENTER"
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
    
    #new column names
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
    
    select_null = pd.isnull(data_all[column_name_operational])
    if not all(abs(data_all.loc[~select_null,column_name_spot]-data_all.loc[~select_null,column_name_operational]-data_all.loc[~select_null,column_name_excess])<1):
        print("***ERROR: spot != operational + excess")
    data_all.loc[select_null,column_name_excess]=data_all.loc[select_null,column_name_spot]
    
    #rename legal entity's names
    data_all.loc[~data_all[column_name_company].isin([2001,2002,2004,2006,2011,2014,2021,2022,2023,2026,2028,2029,2613]),column_name_company]="Others"
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
    data_all.loc[["GmbH" in x for x in data_all[column_name_company]],column_name_company] = 'SSBT GmbH'
    
    #rename product types
    prod_type_mapping = {
        "INTEREST BEARING FOREIGN DEPOSITS":"IBDDA",\
        "INTEREST BEARING DOMESTIC DEPOSITS":"IBDDA",\
        "DEMAND DEPOSITS":"DDA"}
    
    data_all[column_name_product_type] = [prod_type_mapping[x] if x in prod_type_mapping else x for x in data_all[column_name_product_type]]
    data_all[column_name_product_subtype] = [x.replace(" BRANCH","") if isinstance(x,str) else x for x in data_all[column_name_product_subtype]]
    data_all["ProductTypes"] = data_all[column_name_product_type]+"-"+data_all[column_name_product_subtype]

    # behavior group
    data_all.loc[data_all[column_name_parent]=="UnregulatedFund",column_name_group]="Excluded Funds"
    data_all.loc[(pd.isnull(data_all[column_name_group])) & (data_all[column_name_parent]!="UnregulatedFund"),column_name_group]="Funds Miss Model Info"
    data_all.loc[(~pd.isnull(data_all[column_name_group])) & (data_all[column_name_parent]!="UnregulatedFund") & (data_all["ND_IND"]=="Y"),column_name_group]="Non-Discretionary Funds"
    data_all.loc[data_all[column_name_group]=='INTRA-DAY',column_name_group] = "Intraday Clients"
    data_all.loc[data_all[column_name_group]=='Multi-Day Low',column_name_group] = "Multi-Day Low Clients"
    data_all.loc[data_all[column_name_group]=='Multi-Day High',column_name_group] = "Multi-Day High Clients"
    
    # region
    data_all.loc[data_all[column_name_parent]=="UnregulatedFund",column_name_region]="Excluded Funds"
    data_all.loc[(pd.isnull(data_all[column_name_region])) & (data_all[column_name_parent]!="UnregulatedFund"),column_name_region]="Funds Miss Region Info"

    # Currency
    data_all.loc[[x not in ["USD","AUD","CAD","CHF","EUR","GBP","JPY"] for x in data_all[column_name_currency]],column_name_currency] = "Others"
    
    # product types
    IBDDADomestic = ['IBDDA-MONEY MARKET','IBDDA-NOW ACCOUNTS','IBDDA-EXTERNAL']
    data_all.loc[[x in IBDDADomestic for x in data_all['ProductTypes']],'ProductTypes'] = "IBDDA-DOMESTIC"
    data_all.loc[[x not in ["DDA-DOMESTIC","IBDDA-CAYMAN","IBDDA-DOMESTIC","IBDDA-EUROPEAN","IBDDA-LONDON"] for x in data_all['ProductTypes']],'ProductTypes'] = "Others"

    # country of domicile
    data_all.loc[[x not in ['US','UNK','LU','IE','DK','GB','DE','AU','CA','CN','KY','JP'] for x in data_all['COUNTRY_DOMICILE']],'COUNTRY_DOMICILE'] = "Other"
    
    data_all.loc[["MONEY" in x for x in data_all[column_name_style].fillna("_Missing")],column_name_stylegroup] = "CASH - MONEY MARKETS"
    Alternatives = ['CORPORATE - GENERAL','UNKNOWN_STT','INCOME MIXED','REAL ESTATE INVESTMENTS','OTHER_PWC','DERIVATIVES','PRIVATE PLACEMENTS']
    data_all.loc[[x in Alternatives for x in data_all[column_name_stylegroup].fillna("_Missing")],column_name_stylegroup] = "ALTERNATIVE"
    

    return(data_all)