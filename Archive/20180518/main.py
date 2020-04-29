# -*- coding: utf-8 -*-
"""
Created on Sun Nov 19 20:38:45 2017
Run Operational Deposit Model and Reports
@author: e620927
"""
PRODUCTION_ENVRIONMENT = True
import os
import gc

if PRODUCTION_ENVRIONMENT:
    code_dir = "Z:/Charles/ORM/SourceCodes-Production"
else:    
    code_dir = "Z:/Charles/ORM/SourceCodes"
os.chdir(code_dir)

import Process_TRANS_PROD_TREASURY_COMMON_DEPOSITS
gc.collect()

import Process_FX_Rate
gc.collect()

import Process_Balance_By_Currency
gc.collect()

import main_TRANS_PROD_TREASURY_COMMON_DEPOSITS
gc.collect()

