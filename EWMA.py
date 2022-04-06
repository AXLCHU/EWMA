import math
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
from scipy.stats import norm
import time

ZC = pd.read_excel('C:\\***.xlsx', sheet_name='ZC')

holding_period = 1

ZC['ret_FR_2Y'] = ZC['FR_2Y'].diff(periods = holding_period)

eof_zc = len(ZC)
lookback_period = 250
hist_seed = 60
decay = 0.94
histo = 2500
confidence_level = 0.99
nb_ES_values = int(lookback_period*(1 - confidence_level))

VaR = pd.DataFrame(columns=['Date','VaR'])
VaR = []

for i in range(histo):
    ZC.loc[ZC.index[eof_zc - lookback_period -i],'vol_FR_2Y'] = np.std(ZC['ret_FR_2Y'][eof_zc - lookback_period - hist_seed + 1 - i: eof_zc - lookback_period - i])
    for t in range(eof_zc - lookback_period + 1 - i, eof_zc - i):
        ZC.loc[t,'vol_FR_2Y'] = np.sqrt(decay * ZC.loc[t-1,'vol_FR_2Y']**2 + (1-decay) * ZC.loc[t-1,'ret_FR_2Y']**2)
        ZC.loc[t,'resc_ret_FR_2Y'] = ZC.loc[t,'ret_FR_2Y'] * ((ZC.loc[t,'vol_FR_2Y'] + ZC.loc[eof_zc - i - 1,'vol_FR_2Y']) / (2*ZC.loc[t,'vol_FR_2Y'])) # MID
    Worst_CVaR = ZC.loc[eof_zc - lookback_period - i : eof_zc - i , :].nsmallest(nb_ES_values, ['resc_ret_FR_2Y'])
    Worst_CVaR_histo = ZC.loc[eof_zc - lookback_period - i : eof_zc - i , :].nsmallest(nb_ES_values, ['ret_FR_2Y'])
    VaR.append({'Date': ZC.loc[eof_zc - i - 1, 'Date'], 
                'CVaR': np.mean(Worst_CVaR['resc_ret_FR_2Y']), 
                'VaR': -np.quantile(ZC.loc[eof_zc - lookback_period - i:eof_zc - i,'resc_ret_FR_2Y'], confidence_level,axis=0),
               'VaR_histo': -np.quantile(ZC.loc[eof_zc - lookback_period - i:eof_zc - i,'ret_FR_2Y'], confidence_level,axis=0),
                'CVaR_histo': np.mean(Worst_CVaR_histo['ret_FR_2Y'])
               })

