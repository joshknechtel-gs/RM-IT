'''
Most of the CLO Optimizer Tools
within python use>> import CLOutils as clo
function calls from notebook are then 
e.g. clo.moodys_adjusted_warf 
within excel the functions should be in 
list of functions
'''
import numpy as np  
import pandas as pd
import datetime as dt
#import blpapi
#from xbbg import blp
from pyxll import xl_func

path = 'Z:/Shared/Risk Management and Investment Technology/CLO Optimization/'
file = 'CLO17 portfolio as of 04.15.21.xlsm'
filepath= path+file
CLO_tab = 'CLO 17 Port as of 4.15'
Bid_tab = 'Bid.Ask 4.15'

# Default trade size if not specified
pot_trade_size=1e6

############################################################
### Key Constraints & Tests
############################################################
@xl_func("dataframe<index=True>, dataframe<index=True>,dataframe<index=True>:  dataframe<index=True>")
def moodys_adjusted_warf(df,moodys_score,moodys_rfTable):
    """
    This function creates the new Moody's Ratings Factor based 
    on the old Moody's rating.
    
    Arg in:
        df: the input data frame (from the MASTER table d/l'd from BMS)
        moodys_score: dataframe with alphanumeric rating to numeric map (1 to 1 map; linear)
        moodys_rfTable: dataframe with alphanumeric rating to new WARF numeric (1 to 1 map; 1 to 1000 values)
    """
    score = df['Moodys CFR'].map(dict(moodys_score[['Moodys','Score']].values))
    updown = df['Moodys Issuer Watch'].\
        apply(lambda x: -1 if x == 'Possible Upgrade' else 1 if x == 'Possible Downgrade' else 0)
    aScore = score + updown
    Adjusted_CFR_for_WARF = aScore.map(dict(moodys_score[['Score','Moodys']].values))
    # I keep the same column name as Jeff to make it easier to double check values
    df['Adj. WARF NEW'] = Adjusted_CFR_for_WARF.map(dict(moodys_rfTable[['Moody\'s Rating Factor Table','Unnamed: 10']].values))
    return df
################################################################
@xl_func
def sp_recovery_rate(model_df,lien,new_rr,bond_table):
    """
    This function get the S&P recovery rate as a percent. If it doesn't exist
    in the master field, it will look up in the appropriate first and second 
    lien tables, if not, will look up the bond table.
    
    Arg in:
        model_df: the input data frame (from the MASTER table d/l'd from BMS)
        lien: a DF table with the RR's for first and second lien by country
        new_rr: a df mapping of the old notation for RR to a new RR in percentage
        bond_table: split out of a table for RR for bonds
    Arg out:
        model_df with inserted new column 'S&P Recovery Rate (AAA)'
    """
     
    # if it the Recovery rate exists lookup in AAA table
    model_df['S&P Recovery Rate (AAA)'] = model_df['S&P Recovery'].\
        map(dict(new_rr[['S&P Recovery Rating\nand Recovery\nIndicator of\nCollateral Obligations','“AAA”']].values))
    
    # doesn't exist, but first lien, use first lien table
    model_df.loc[pd.isna(model_df['S&P Recovery']) & (model_df['Lien Type']== 'First Lien'),'S&P Recovery Rate (AAA)'] =\
        model_df.loc[pd.isna(model_df['S&P Recovery']) & (model_df['Lien Type']== 'First Lien'),'Issuer Country'].\
        map(dict(lien[['Country Abv','RR']].values))
    
    
    # doesn't exist, but 2nd lien, use 2nd lien table
    model_df.loc[pd.isna(model_df['S&P Recovery']) & (model_df['Lien Type']== 'Second Lien'),'S&P Recovery Rate (AAA)'] = \
        model_df.loc[pd.isna(model_df['S&P Recovery']) & (model_df['Lien Type']== 'Second Lien'),'Issuer Country'].\
        map(dict(lien[['Country Abv','RR.2nd']].values))
    
    # the bonds
    model_df.loc[pd.isna(model_df['S&P Recovery']) & pd.isna(model_df['Lien Type']),'S&P Recovery Rate (AAA)'] = \
        model_df.loc[pd.isna(model_df['S&P Recovery']) & pd.isna(model_df['Lien Type']),'Issuer Country'].\
        map(dict(bond_table[['Country Abv.1','RR.1']].values))

    return model_df
################################################################
@xl_func
def diversity_score(model_df, weight_col='Par_no_default'):
    """
    This function calculates the Moody's Industry Diversity Score for the CLO
    
    Arg in:
        model_df: the input data frame (from the MASTER table d/l'd from BMS)
        #ind_avg_eu: Moody's discrete lookup table that maps AIEUS to IDS, need to be sorted
    Arg out:
        dscore: the scalar measure of the IDS
    """
    # calling this each time time (like in MC Diversity)
    # makes the function unbearably slow. I tried setting it
    # as a global variable but that didn't solve the issue
    # I need to look into a solution to set something up 
    # permanently in memory after the first call (though global would do it)
    # ind_avg_eu = get_ind_avg_eu_table(filepath,sheet='Diversity')
    # thus I've hard-coded it, but if this ever changes from Moody's
    # we need to update here.
    ind_avg_eu = pd.DataFrame([[ 0.    ,  0.    ], [ 0.05  ,  0.1   ], [ 0.15  ,  0.2   ],
       [ 0.25  ,  0.3   ], [ 0.35  ,  0.4   ], [ 0.45  ,  0.5   ], [ 0.55  ,  0.6   ], 
       [ 0.65  ,  0.7   ], [ 0.75  ,  0.8   ], [ 0.85  ,  0.9   ], [ 0.95  ,  1.    ],
       [ 1.05  ,  1.05  ], [ 1.15  ,  1.1   ], [ 1.25  ,  1.15  ], [ 1.35  ,  1.2   ],
       [ 1.45  ,  1.25  ], [ 1.55  ,  1.3   ], [ 1.65  ,  1.35  ], [ 1.75  ,  1.4   ],
       [ 1.85  ,  1.45  ], [ 1.95  ,  1.5   ], [ 2.05  ,  1.55  ], [ 2.15  ,  1.6   ],
       [ 2.25  ,  1.65  ], [ 2.35  ,  1.7   ], [ 2.45  ,  1.75  ], [ 2.55  ,  1.8   ],
       [ 2.65  ,  1.85  ], [ 2.75  ,  1.9   ], [ 2.85  ,  1.95  ], [ 2.95  ,  2.    ],
       [ 3.05  ,  2.0333], [ 3.15  ,  2.0667], [ 3.25  ,  2.1   ], [ 3.35  ,  2.1333],
       [ 3.45  ,  2.1667], [ 3.55  ,  2.2   ], [ 3.65  ,  2.2333], [ 3.75  ,  2.2667],
       [ 3.85  ,  2.3   ], [ 3.95  ,  2.3333], [ 4.05  ,  2.3667], [ 4.15  ,  2.4   ],
       [ 4.25  ,  2.4333], [ 4.35  ,  2.4667], [ 4.45  ,  2.5   ], [ 4.55  ,  2.5333],
       [ 4.65  ,  2.5667], [ 4.75  ,  2.6   ], [ 4.85  ,  2.6333], [ 4.95  ,  2.6667],
       [ 5.05  ,  2.7   ], [ 5.15  ,  2.7333], [ 5.25  ,  2.7667], [ 5.35  ,  2.8   ],
       [ 5.45  ,  2.8333], [ 5.55  ,  2.8667], [ 5.65  ,  2.9   ], [ 5.75  ,  2.9333],
       [ 5.85  ,  2.9667], [ 5.95  ,  3.    ], [ 6.05  ,  3.025 ], [ 6.15  ,  3.05  ],
       [ 6.25  ,  3.075 ], [ 6.35  ,  3.1   ], [ 6.45  ,  3.125 ], [ 6.55  ,  3.15  ],
       [ 6.65  ,  3.175 ], [ 6.75  ,  3.2   ], [ 6.85  ,  3.225 ], [ 6.95  ,  3.25  ],
       [ 7.05  ,  3.275 ], [ 7.15  ,  3.3   ], [ 7.25  ,  3.325 ], [ 7.35  ,  3.35  ],
       [ 7.45  ,  3.375 ], [ 7.55  ,  3.4   ], [ 7.65  ,  3.425 ], [ 7.75  ,  3.45  ],
       [ 7.85  ,  3.475 ], [ 7.95  ,  3.5   ], [ 8.05  ,  3.525 ], [ 8.15  ,  3.55  ],
       [ 8.25  ,  3.575 ], [ 8.35  ,  3.6   ], [ 8.45  ,  3.625 ], [ 8.55  ,  3.65  ],
       [ 8.65  ,  3.675 ], [ 8.75  ,  3.7   ], [ 8.85  ,  3.725 ], [ 8.95  ,  3.75  ],
       [ 9.05  ,  3.775 ], [ 9.15  ,  3.8   ], [ 9.25  ,  3.825 ], [ 9.35  ,  3.85  ],
       [ 9.45  ,  3.875 ], [ 9.55  ,  3.9   ], [ 9.65  ,  3.925 ], [ 9.75  ,  3.95  ],
       [ 9.85  ,  3.975 ], [ 9.95  ,  4.    ], [10.05  ,  4.01  ], [10.15  ,  4.02  ],
       [10.25  ,  4.03  ], [10.35  ,  4.04  ], [10.45  ,  4.05  ], [10.55  ,  4.06  ],
       [10.65  ,  4.07  ], [10.75  ,  4.08  ], [10.85  ,  4.09  ], [10.95  ,  4.1   ],
       [11.05  ,  4.11  ], [11.15  ,  4.12  ], [11.25  ,  4.13  ], [11.35  ,  4.14  ],
       [11.45  ,  4.15  ], [11.55  ,  4.16  ], [11.65  ,  4.17  ], [11.75  ,  4.18  ],
       [11.85  ,  4.19  ], [11.95  ,  4.2   ], [12.05  ,  4.21  ], [12.15  ,  4.22  ],
       [12.25  ,  4.23  ], [12.35  ,  4.24  ], [12.45  ,  4.25  ], [12.55  ,  4.26  ],
       [12.65  ,  4.27  ], [12.75  ,  4.28  ], [12.85  ,  4.29  ], [12.95  ,  4.3   ],
       [13.05  ,  4.31  ], [13.15  ,  4.32  ], [13.25  ,  4.33  ], [13.35  ,  4.34  ],
       [13.45  ,  4.35  ], [13.55  ,  4.36  ], [13.65  ,  4.37  ], [13.75  ,  4.38  ],
       [13.85  ,  4.39  ], [13.95  ,  4.4   ], [14.05  ,  4.41  ], [14.15  ,  4.42  ],
       [14.25  ,  4.43  ], [14.35  ,  4.44  ], [14.45  ,  4.45  ], [14.55  ,  4.46  ],
       [14.65  ,  4.47  ], [14.75  ,  4.48  ], [14.85  ,  4.49  ], [14.95  ,  4.5   ],
       [15.05  ,  4.51  ], [15.15  ,  4.52  ], [15.25  ,  4.53  ], [15.35  ,  4.54  ],
       [15.45  ,  4.55  ], [15.55  ,  4.56  ], [15.65  ,  4.57  ], [15.75  ,  4.58  ],
       [15.85  ,  4.59  ], [15.95  ,  4.6   ], [16.05  ,  4.61  ], [16.15  ,  4.62  ],
       [16.25  ,  4.63  ], [16.35  ,  4.64  ], [16.45  ,  4.65  ], [16.55  ,  4.66  ],
       [16.65  ,  4.67  ], [16.75  ,  4.68  ], [16.85  ,  4.69  ], [16.95  ,  4.7   ],
       [17.05  ,  4.71  ], [17.15  ,  4.72  ], [17.25  ,  4.73  ], [17.35  ,  4.74  ],
       [17.45  ,  4.75  ], [17.55  ,  4.76  ], [17.65  ,  4.77  ], [17.75  ,  4.78  ],
       [17.85  ,  4.79  ], [17.95  ,  4.8   ], [18.05  ,  4.81  ], [18.15  ,  4.82  ],
       [18.25  ,  4.83  ], [18.35  ,  4.84  ], [18.45  ,  4.85  ], [18.55  ,  4.86  ],
       [18.65  ,  4.87  ], [18.75  ,  4.88  ], [18.85  ,  4.89  ], [18.95  ,  4.9   ],
       [19.05  ,  4.91  ], [19.15  ,  4.92  ], [19.25  ,  4.93  ], [19.35  ,  4.94  ], [19.45  ,  4.95  ],
       [19.55  ,  4.96  ], [19.65  ,  4.97  ], [19.75  ,  4.98  ], [19.85  ,  4.99  ], [19.95  ,  5.    ]], 
                              columns = ['Aggregate\nIndustry\nEquivalent\nUnit Score', 'Industry\nDiversity\nScore'])    
    #first create the Par amount filtering out defaults
    #model_df['Par_no_default'] = model_df['Total']
    #model_df.loc[model_df['Default']=='Y','Par_no_default'] = 0
    div_df = model_df[['Parent Company','Moodys Industry',weight_col]].copy()
    div_df.sort_values(by='Moodys Industry',inplace=True)

    # this keeps the industry, but groups on parent company for multiple loans
    test = div_df.groupby(by=['Parent Company','Moodys Industry']).sum()
    avg_par_amt = test.sum()/test.count()   
    
    # create the EU score for each parent
    # Lesser of 1 and Issuer Par Amount for such issuer divided by the Average Par Amount.
    test['EU'] = test[[weight_col]]/test[[weight_col]].mean()
    test.loc[test['EU']>1,'EU']=1
    
    # groupby Industry for the Ind Div Score
    IDS = test.groupby(by=['Moodys Industry']).sum()

    # this is like vlookup(..,TRUE) where the nearest match on merge is used, direction controls how
    # backward is the lesser if EU falls between AIEUS marks
    df_merged = pd.merge_asof(IDS.sort_values('EU'), ind_avg_eu, left_on='EU', 
                          right_on='Aggregate\nIndustry\nEquivalent\nUnit Score', direction='backward', suffixes=['', '_2'])
    dscore = df_merged['Industry\nDiversity\nScore'].sum()
    return dscore
################################################################
@xl_func
def weighted_average(model_df,cols):
    """
    Calculates the weighted average variable
    Args in:
        model_df
        cols: must be a list of two elements with the first element
                the weighting elements and the second the statistic to be weighted
                i.e. ['weight','stat']
    Arg out: 
        wa: scalar
    """
    #wa = model_df[cols].apply(lambda x: (x[0]*x[1]).sum()/x[0].sum())
    wa = (model_df[cols[0]]*model_df[cols[1]]).sum()/model_df[cols[0]].sum()
    return wa

# this were all based off Total, but should be PnD or PnD_postTrade 
# former weight_col='Total', now 
################################################################
@xl_func
def percentage_C(model_df,weight_col='Par_no_default'):  
    perC = model_df.loc[model_df['Adjusted CFR for WARF'].str.match('C'),[weight_col]].sum()/model_df[[weight_col]].sum()
    return perC.values[0]
################################################################
@xl_func
def percentage_SecondLien(model_df,weight_col='Par_no_default'):  
    perC = model_df.loc[model_df['Lien Type']=='Second Lien',[weight_col]].sum()/model_df[[weight_col]].sum()
    return perC.values[0]
################################################################
@xl_func
def percentage_SubEighty(model_df,weight_col='Par_no_default'):  
    perC = model_df.loc[model_df['Blended Price']<80,[weight_col]].sum()/model_df[[weight_col]].sum()
    return perC.values[0]
################################################################
@xl_func
def percentage_SubNinety(model_df,weight_col='Par_no_default'):  
    perC = model_df.loc[model_df['Blended Price']<90,[weight_col]].sum()/model_df[[weight_col]].sum()
    return perC.values[0]
################################################################
@xl_func
def percentage_CovLite(model_df,weight_col='Par_no_default'):  
    perC = model_df.loc[model_df['Cov Lite']=='Yes',[weight_col]].sum()/model_df[[weight_col]].sum()
    return perC.values[0]
################################################################
@xl_func
def BAPP(model_df):
    model_df['Blended Actual Purchase Prices'] = \
        model_df[['Addtl Purchase Amt','Purch Price of Addtl Purch',
                    'Current Portfolio','Actual Purch Price of Current Positions']].\
        apply(lambda x: 100*((x[0]*x[1]/100)+(x[2]*x[3]/100))/(x[0]+x[2]),axis=1) #I get runtime warnings often
    model_df.loc[model_df['Blended Actual Purchase Prices'].isna(),'Blended Actual Purchase Prices'] = 0
    return model_df
################################################################
@xl_func
def blended_price(model_df):
    model_df['Blended Price'] = model_df[['Potential Trades',
            'Addtl Purchase Amt','Blended Actual Purchase Prices','Total','Bid','Ask','Current Portfolio']].\
            apply(lambda x: ((x[0]*x[4]/100+(x[6]+x[1])*x[2])/x[3])*100 if x[0]<1 else \
                            ((x[0]*x[5]/100+(x[6]+x[1])*x[2])/x[3])*100,axis=1 )  #I get runtime warnings often
    model_df.loc[model_df['Blended Price'].isna(),'Blended Price'] = 0
    return model_df
################################################################
#
@xl_func
def par_build_loss(model_df,pot_trades='Potential Trades'):
    model_df['Par_Build_Loss_Sale'] = model_df[[pot_trades,
                                                    'Bid','Actual Purch Price of Current Positions']].\
            apply(lambda x: ((-x[0]*x[1]/100)-(-x[0]*x[2]/100) if x[0]<0 else 0),axis=1)
    model_df['Par_Build_Loss_Buy'] = model_df[[pot_trades,'Ask']].\
            apply(lambda x: ((-x[0]*x[1]/100)+x[0]) if x[0]>0 else 0,axis=1)
    model_df['Total_Par_Build_Loss'] = model_df[['Par_Build_Loss_Sale','Par_Build_Loss_Buy']].sum(axis=1)
    return model_df
################################################################
@xl_func
def par_burn_new(model_df,pot_trades='Potential Trades'):
    model_df['Par_Build_Loss_Sale'] = model_df[[pot_trades,
                                                    'Bid','Actual Purch Price of Current Positions']].\
            apply(lambda x: ((-x[0]*x[1]/100)-(-x[0]*100/100) if x[0]<0 else 0),axis=1)
    model_df['Par_Build_Loss_Buy'] = model_df[[pot_trades,'Ask']].\
            apply(lambda x: ((-x[0]*x[1]/100)+x[0]) if x[0]>0 else 0,axis=1)
    model_df['Total_Par_Build_Loss'] = model_df[['Par_Build_Loss_Sale','Par_Build_Loss_Buy']].sum(axis=1)
    return model_df
################################################################
@xl_func
def mil_par_build_loss(model_df,pot_trade_size=1e6):
    """
    This function calculates the Par B/L Sale, Buy and Total for ALL
    loans assuming 1mn notional (optional change size).  Essentially,
    consider this akin to Marginal Contribution to Par B/L
    
    Arg in:
        model_df
        pot_trade_size (default=100000)
    Arg out:
        model_df with additional columns, Mil_Par_BL_Sale, Mil_Par_BL_Buy, Mil_Par_Build_Loss
    """
    # you can buy anything but you can only sell what you own!  Changed to min pot_trade_size and current
    trade = -pot_trade_size
    model_df.loc[model_df['Current Portfolio']>0,['Mil_Par_BL_Sale']] = \
        model_df.loc[model_df['Current Portfolio']>0,['Current Portfolio','Bid','Actual Purch Price of Current Positions']].\
        apply(lambda x: ((-min(trade,x[0])*x[1]/100)-(-min(trade,x[0])*x[2])),axis=1)  # I think this is wrong in xls
    trade = pot_trade_size
    model_df['Mil_Par_BL_Buy'] = model_df[['Ask']].\
            apply(lambda x: ((-trade*x/100)+trade),axis=1)
    model_df['Mil_Par_Build_Loss'] = model_df[['Mil_Par_BL_Sale','Mil_Par_BL_Buy']].sum(axis=1)
    return model_df
################################################################
@xl_func
def mil_parburn_new(model_df,pot_trade_size=1e6):
    """
    This function calculates the Par B/L Sale, Buy and Total for ALL
    loans assuming 1mn notional (optional change size).  Essentially,
    consider this akin to Marginal Contribution to Par B/L
    
    Arg in:
        model_df
        pot_trade_size (default=100000)
    Arg out:
        model_df with additional columns, Mil_Par_BL_Sale, Mil_Par_BL_Buy, Mil_Par_Build_Loss
    """
    # you can buy anything but you can only sell what you own!  Changed to min pot_trade_size and current
    trade = -pot_trade_size
    model_df.loc[model_df['Current Portfolio']>0,['Mil_Par_BL_Sale']] = \
        model_df.loc[model_df['Current Portfolio']>0,['Current Portfolio','Bid']].\
        apply(lambda x: ((-min(trade,x[0])*x[1]/100)-(-min(trade,x[0]))),axis=1)  
    #,'Actual Purch Price of Current Positions'
    trade = pot_trade_size
    model_df['Mil_Par_BL_Buy'] = model_df[['Ask']].\
            apply(lambda x: ((-trade*x/100)+trade),axis=1)
    model_df['Mil_Par_Build_Loss'] = model_df[['Mil_Par_BL_Sale','Mil_Par_BL_Buy']].sum(axis=1)
    return model_df
################################################################
@xl_func
def mc_WARF(model_df,pot_trade_size=1e6):
    oldWARF = weighted_average(model_df,cols=['Par_no_default','Adj. WARF NEW'])
    
    for row in model_df.index:
        div_df = model_df[['Par_no_default','Adj. WARF NEW']].copy()    
        div_df.loc[row,'Par_no_default'] += pot_trade_size
        model_df.loc[row,'MC WARF'] = weighted_average(div_df,cols=['Par_no_default','Adj. WARF NEW']) - oldWARF

    return model_df
################################################################
@xl_func
def mc_WAS(model_df,pot_trade_size=1000000):
    oldWARF = weighted_average(model_df,cols=['Par_no_default','Spread'])
  
    for row in model_df.index:
        div_df = model_df[['Par_no_default','Spread']].copy()    
        div_df.loc[row,'Par_no_default'] += pot_trade_size
        model_df.loc[row,'MC WAS'] = weighted_average(div_df,cols=['Par_no_default','Spread']) - oldWARF

    return model_df
################################################################
@xl_func
def mc_WAPP(model_df,pot_trade_size=1000000):
    oldWARF = weighted_average(model_df,cols=['Par_no_default','Ask'])

    for row in model_df.index:
        div_df = model_df[['Par_no_default','Ask']].copy()    
        div_df.loc[row,'Par_no_default'] += pot_trade_size
        model_df.loc[row,'MC WAPP'] = weighted_average(div_df,cols=['Par_no_default','Ask']) - oldWARF
    return model_df
################################################################
@xl_func
def MC_diversity_score(model_df, pot_trade_size=1000000):
    """
    This function calculates the Marginal Contribution Moody's Industry Diversity Score 
    for all loan given a potential trade size (default 1mn)
    
    Arg in:
        model_df: the input data frame (from the MASTER table d/l'd from BMS)
        #ind_avg_eu: Moody's discrete lookup table that maps AIEUS to IDS, need to be sorted
        pot_trade_size
    Arg out:
        model_df: with added MC Div Score field
    """
    curr_DS = diversity_score(model_df)
    for row in model_df.index:
        div_df = model_df[['Parent Company','Moodys Industry','Par_no_default']].copy()    
        div_df.loc[row,'Par_no_default'] += pot_trade_size
        model_df.loc[row,'MC Div Score'] = diversity_score(div_df) - curr_DS
    return model_df
################################################################
@xl_func
def create_marginal_stats(model_df, pot_trade_size=1e6):
    model_df = mil_parburn_new(model_df,pot_trade_size)
    model_df = mc_WARF(model_df,pot_trade_size)
    model_df = MC_diversity_score(model_df, pot_trade_size)
    model_df = mc_WAS(model_df,pot_trade_size)
    model_df = mc_WAPP(model_df,pot_trade_size)
    return model_df
################################################################
###  Main Port Stats Function
################################################################
@xl_func("dataframe<index=True>, str: dataframe<index=True>", auto_resize=True)
def Port_stats(model_df, weight_col='Par_no_default'):
    """
    Arg in:
        model_df
        #ind_avg_eu: table (df) Moody's discrete lookup table that maps AIEUS to IDS
    
    - Estimated Libor
    - Minimum Floating Spread Test - Without Libor Floors
    - Minimum Floating Spread Test - WithLibor Floors (adj. All in Rate)
    - Maximum Moody's Rating Factor Test
    - Maximum Moody's Rating Factor Test (NEW WARF)
    - Maximum Moody's Rating Factor Test (Orig WARF)
    - Minimum Weighted Average Moody's Recovery Rate Test
    - Minimum Weighted Average S&P Recovery Rate Class A-1a
    - Moody's Diversity Test
    - WAP (Current Positions use Actual purchase price, all others use Ask price)
    - Total Portfolio Par (excluding Defaults)
    - Total Portfolio Par
    - Current Portfolio 
    
    
    """
    
    Port_stats_df = pd.DataFrame(np.nan,index=['Min Floating Spread Test - no Libor Floors',
        'Min Floating Spread Test - With Libor Floors',
        'Max Moodys Rating Factor Test (NEW WARF)',
        'Max Moodys Rating Factor Test (Orig WARF)',
        'Min Moodys Recovery Rate Test',
        'Min S&P Recovery Rate Class A-1a',
        'Moodys Diversity Test',
        'WAP',
        'Percent C',
        'Percent 2nd Lien',
        'Percent Sub80',
        'Percent Sub90',
        'Percent CovLite',
    #    'Pot Par B/L Sale',
    #    'Pot Par B/L Buy',
    #    'Pot Par B/L Total',
        'Total Portfolio Par (excl. Defaults)',
    #    'Total Portfolio Par',
     #   'Current Portfolio'
                                              ],columns = ['Portfolio Stats'])

    # Par_no_default works for current positions only
    # if trying to compare potential new trades, need updated field
    # also Total isn't dynamically updated for Potential Trades once
    # model_df is created.  Need to find solution
    
    Port_stats_df.loc['Min Floating Spread Test - no Libor Floors','Portfolio Stats'] = \
        weighted_average(model_df,cols=[weight_col,'Spread'])*100
    Port_stats_df.loc['Min Floating Spread Test - no Libor Floors'] = \
        Port_stats_df.loc['Min Floating Spread Test - no Libor Floors'].apply('{:.2f}%'.format)
    
    Port_stats_df.loc['Min Floating Spread Test - With Libor Floors','Portfolio Stats'] = \
        weighted_average(model_df,cols=[weight_col,'Adj. All in Rate'])*100
    Port_stats_df.loc['Min Floating Spread Test - With Libor Floors'] = \
        Port_stats_df.loc['Min Floating Spread Test - With Libor Floors'].apply('{:.2f}%'.format)
    
    Port_stats_df.loc['Max Moodys Rating Factor Test (NEW WARF)','Portfolio Stats'] = \
        weighted_average(model_df,cols=[weight_col,'Adj. WARF NEW'])
    Port_stats_df.loc['Max Moodys Rating Factor Test (NEW WARF)'] = \
        Port_stats_df.loc['Max Moodys Rating Factor Test (NEW WARF)'].apply('{:.0f}'.format)
    
    Port_stats_df.loc['Max Moodys Rating Factor Test (Orig WARF)','Portfolio Stats'] = \
        weighted_average(model_df,cols=[weight_col,'WARF'])
    Port_stats_df.loc['Max Moodys Rating Factor Test (Orig WARF)'] = \
        Port_stats_df.loc['Max Moodys Rating Factor Test (Orig WARF)'].apply('{:.0f}'.format)
    
    Port_stats_df.loc['Min Moodys Recovery Rate Test','Portfolio Stats'] = \
        weighted_average(model_df,cols=[weight_col,'Moodys Recovery Rate'])*100    
    Port_stats_df.loc['Min Moodys Recovery Rate Test'] = \
        Port_stats_df.loc['Min Moodys Recovery Rate Test'].apply('{:.1f}%'.format)
    
    Port_stats_df.loc['Min S&P Recovery Rate Class A-1a','Portfolio Stats'] = \
            weighted_average(model_df,cols=[weight_col,'S&P Recovery Rate (AAA)'])*100
    Port_stats_df.loc['Min S&P Recovery Rate Class A-1a'] = \
        Port_stats_df.loc['Min S&P Recovery Rate Class A-1a'].apply('{:.1f}%'.format)
    
    Port_stats_df.loc['Moodys Diversity Test','Portfolio Stats'] = diversity_score(model_df, weight_col)
    Port_stats_df.loc['Moodys Diversity Test'] = \
        Port_stats_df.loc['Moodys Diversity Test'].apply('{:.0f}'.format)
    
    ######   Was Total?
    Port_stats_df.loc['WAP','Portfolio Stats'] = \
        weighted_average(model_df,cols=[weight_col,'Blended Price'])
    Port_stats_df.loc['WAP'] = \
        Port_stats_df.loc['WAP'].apply('${:.2f}'.format)
    
    Port_stats_df.loc['Percent C','Portfolio Stats'] = percentage_C(model_df, weight_col)*100
    Port_stats_df.loc['Percent C'] = \
        Port_stats_df.loc['Percent C'].apply('{:.1f}%'.format)
    
    Port_stats_df.loc['Percent 2nd Lien','Portfolio Stats'] = percentage_SecondLien(model_df, weight_col)*100
    Port_stats_df.loc['Percent 2nd Lien'] = \
        Port_stats_df.loc['Percent 2nd Lien'].apply('{:.1f}%'.format)
    
    Port_stats_df.loc['Percent Sub80','Portfolio Stats'] = percentage_SubEighty(model_df, weight_col)*100
    Port_stats_df.loc['Percent Sub80'] = \
        Port_stats_df.loc['Percent Sub80'].apply('{:.1f}%'.format)
    
    Port_stats_df.loc['Percent Sub90','Portfolio Stats'] = percentage_SubNinety(model_df, weight_col)*100
    Port_stats_df.loc['Percent Sub90'] = \
        Port_stats_df.loc['Percent Sub90'].apply('{:.1f}%'.format)
    
    Port_stats_df.loc['Percent CovLite','Portfolio Stats'] = percentage_CovLite(model_df, weight_col)*100
    Port_stats_df.loc['Percent CovLite'] = \
        Port_stats_df.loc['Percent CovLite'].apply('{:.1f}%'.format)

  
    #Port_stats_df.loc['Pot Par B/L Sale','Portfolio Stats'] = model_df['Par_Build_Loss_Sale'].sum()
    #Port_stats_df.loc['Pot Par B/L Sale'] = \
    #    Port_stats_df.loc['Pot Par B/L Sale'].apply('{:,.0f}'.format)    
    
    #Port_stats_df.loc['Pot Par B/L Buy','Portfolio Stats'] = model_df['Par_Build_Loss_Buy'].sum()
    #Port_stats_df.loc['Pot Par B/L Buy'] = \
    #    Port_stats_df.loc['Pot Par B/L Buy'].apply('{:,.0f}'.format)    
    
    #Port_stats_df.loc['Pot Par B/L Total','Portfolio Stats'] = model_df['Total_Par_Build_Loss'].sum()
    #Port_stats_df.loc['Pot Par B/L Total'] = \
    #    Port_stats_df.loc['Pot Par B/L Total'].apply('{:,.0f}'.format)    

    Port_stats_df.loc['Total Portfolio Par (excl. Defaults)','Portfolio Stats'] = model_df[weight_col].sum()
    Port_stats_df.loc['Total Portfolio Par (excl. Defaults)'] = \
        Port_stats_df.loc['Total Portfolio Par (excl. Defaults)'].apply('{:,.0f}'.format)

    
#    Port_stats_df.loc['Total Portfolio Par','Portfolio Stats'] = model_df['Total'].sum()
#    Port_stats_df.loc['Total Portfolio Par'] = \
#        Port_stats_df.loc['Total Portfolio Par'].apply('{:,.0f}'.format)
    
    # current portfolio is Quantity + Add'l Amount (manual) TBA later
#    Port_stats_df.loc['Current Portfolio','Portfolio Stats'] = model_df[['Addtl Purchase Amt','Current Portfolio']].sum(axis=1).sum()
#    Port_stats_df.loc['Current Portfolio'] = \
#        Port_stats_df.loc['Current Portfolio'].apply('{:,.0f}'.format)
    
    return Port_stats_df
################################################################
@xl_func
def comp_Port_stats(model_df):
    
    
    mask = abs(model_df['Current Portfolio']) > 0
    cstats = Port_stats(model_df.loc[mask], 'Now')
    cstats.rename(columns={'Portfolio Stats':'Current Portfolio'},inplace=True)

    mask = abs(model_df['Potential Trades']) > 0
    pstats = Port_stats(model_df.loc[mask], 'Potential Trades')
    pstats.rename(columns={'Portfolio Stats':'Potential Trades (incl Replines)'},inplace=True)
    
    tstats = Port_stats(model_df, 'Total')
    tstats.rename(columns={'Portfolio Stats':'Total Portfolio (incl Trades)'},inplace=True)
    
    cstats = cstats.join(pstats)
    cstats = cstats.join(tstats)
    
    return cstats
################################################################
@xl_func("dataframe<index=True>,str[] array: dataframe<index=True>", auto_resize = True)
def prepost_Port_stats(model_df, cols):
    
    
    mask = abs(model_df['Current Portfolio']) > 0
    cstats = Port_stats(model_df.loc[mask], cols[0])
    cstats.rename(columns={'Portfolio Stats':'Current Portfolio (pre-trade)'},inplace=True)

    tstats = Port_stats(model_df,  cols[1])
    tstats.rename(columns={'Portfolio Stats':'Post-trade Portfolio'},inplace=True)
    
    cstats = cstats.join(tstats)
    
    return cstats
################################################################
@xl_func
def replines(model_df):
    replines = model_df[model_df['Issuer'].str.match('zz_LXREP')]
    repline_stats_df = pd.DataFrame(np.nan,index=['Amount',
        'WAS',
        'WAPP',
        'WARF New',
        'WARF Orig'],columns = ['Repline Stats'])
    pot_trades = replines['Potential Trades']
    repline_stats_df.loc['Amount','Repline Stats'] = pot_trades.sum()/1000000
    repline_stats_df.loc['WAS','Repline Stats'] = (pot_trades*replines['Spread']).sum()/pot_trades.sum()*100
    repline_stats_df.loc['WAPP','Repline Stats'] = (pot_trades*replines['Ask']).sum()/pot_trades.sum()
    repline_stats_df.loc['WARF New','Repline Stats'] = (pot_trades*replines['Adj. WARF NEW']).sum()/pot_trades.sum()
    repline_stats_df.loc['WARF Orig','Repline Stats'] = (pot_trades*replines['WARF']).sum()/pot_trades.sum()
    
    # just formatting
    repline_stats_df.loc['Amount'] = repline_stats_df.loc['Amount'].apply('${:.1f}'.format)
    repline_stats_df.loc['WAS'] = repline_stats_df.loc['WAS'].apply('{:.2f}%'.format)
    repline_stats_df.loc['WAPP'] = repline_stats_df.loc['WAPP'].apply('{:.2f}'.format)
    repline_stats_df.loc['WARF New'] = repline_stats_df.loc['WARF New'].apply('{:.0f}'.format)
    repline_stats_df.loc['WARF Orig'] = repline_stats_df.loc['WARF Orig'].apply('{:.0f}'.format)
    
    return repline_stats_df
################################################################
# functions for reading the relevant spreadsheet data
# I am making them simple and separated so they can be replaced
# by other bespoke solutions like direct APIs from the source,
# also aides readability and debugging
################################################################
@xl_func
def get_master_df(filepath,sheet='MASTER'):
    master_df = pd.read_excel(filepath,sheet_name=sheet,header=1)
    master_df = master_df.loc[:,~master_df.columns.str.match("Unnamed")]
    master_df.rename(columns={'LoanX ID':'LXID'},inplace=True)
    master_df.set_index('LXID', inplace=True)
    return master_df
################################################################
@xl_func
def get_CLO_df(filepath,sheet='CLO 21 Port as of 3.18'):
    CLO_df = pd.read_excel(filepath,sheet_name=sheet,header=6,usecols='A:K')
    CLO_df.dropna(inplace=True)
    CLO_df.rename(columns={'Cusip or LIN':'LXID'},inplace=True)
    CLO_df.set_index('LXID', inplace=True)
    CLO_df = CLO_df.groupby('LXID').agg({'Quantity':'sum','Cost':'sum',
                                         '/Unit':'mean','Price':'mean','Value':'sum',
                                         'Value.1':'sum','(-Loss)':'sum'})
    # these comments are here, because this is what I used to check
    # his pivot table columns M:0
    # CLO_df[['Cusip or LIN','Quantity','/Unit']].sort_values(by='.Cusip or LIN')
    # CLO_df[['Quantity']].sum()  # verified sum of Quantity
    # CLO_df[['/Unit']].mean()    # verified for average /Unit
    return CLO_df
################################################################
@xl_func
def get_bidask_df(filepath,sheet='Bid.Ask 3.18'):
    bidask_df = pd.read_excel(filepath,sheet_name=sheet,header=0)
    bidask_df = bidask_df.loc[:,~bidask_df.columns.str.match("Unnamed")]
    bidask_df.set_index('LXID', inplace=True)
    return bidask_df
################################################################
@xl_func
def get_moodys_rating2rf_tables(filepath,sheet='New WARF'):
    moodys_score = pd.read_excel(filepath,sheet_name=sheet,header=0,usecols='E:F')
    moodys_rfTable = pd.read_excel(filepath,sheet_name=sheet,header=0,usecols='J:K')
    return moodys_score, moodys_rfTable
################################################################
@xl_func
def get_recovery_rate_tables(filepath,sheet='SP RR Updated'):
    new_sp_rr = pd.read_excel(filepath, sheet_name=sheet, header=1, usecols='L:M')
    new_sp_rr.dropna(how='all',inplace=True)

    lien_rr = pd.read_excel(filepath, sheet_name=sheet, header=1, usecols='A:I')
    lien_rr.dropna(how='all',inplace=True)

    bond_split = lien_rr[lien_rr['Country.1']=='Bonds'].index.values[0]
    bond_table = lien_rr.loc[bond_split+1:]
    lien_rr = lien_rr.loc[:bond_split-1]
    lien_rr.drop(columns=['Unnamed: 4','Country Abv.1','Country.1','Group.1'],inplace=True)
    lien_rr.rename(columns={'RR.1':'RR.2nd'},inplace=True)
    
    return new_sp_rr, lien_rr, bond_table
################################################################
@xl_func("str, str: object",auto_resize=True)
def get_ind_avg_eu_table(filepath,sheet='Diversity'):
    ind_avg_eu = pd.read_excel(filepath, sheet_name=sheet, header=8, usecols='K:L')
    ind_avg_eu.dropna(how='all',inplace=True)
    return ind_avg_eu
################################################################
@xl_func
def get_pot_trades(filepath,sheet='Model Portfolio'):
    pot_trades = pd.read_excel(filepath,sheet_name=sheet,header=15,usecols='C:G')
    pot_trades.rename(columns={'LX ID':'LXID'},inplace=True)
    pot_trades.set_index('LXID', inplace=True)

    return pot_trades
################################################################
@xl_func
def model_pricing(model_df):
    model_df.loc[model_df['Close Offer'].isna(),'Close Offer'] = 99
    #model_df.loc[model_df['Issuer'].str.match('zz_LXREP'),'Close Offer'] = 99.5
    #model_df.loc[model_df['Issuer'].str.match('zz_LXREP03'),'Close Offer'] = 99
    #model_df.loc[model_df['Issuer'].str.match('zz_LXREP12'),'Close Offer'] = 99
    model_df.loc[model_df['Close Bid'].isna(),'Close Bid'] = 99
    model_df.loc[model_df['/Unit'].isna(),'/Unit'] = 0
    model_df.rename(columns={'Close Offer':'Ask','Close Bid':'Bid',
                               '/Unit':'Actual Purch Price of Current Positions'},inplace=True)
    model_df.loc[model_df['Actual Purch Price of Current Positions'].isna(),'Actual Purch Price of Current Positions'] = 0
    model_df['Current Portfolio'] = model_df['Quantity']
    model_df.loc[model_df['Current Portfolio'].isna(),'Current Portfolio'] = 0
    model_df.loc[model_df['Potential Trades'].isna(),'Potential Trades'] = 0
    model_df.loc[model_df['Addtl Purchase Amt'].isna(),'Addtl Purchase Amt'] = 0
    model_df.loc[model_df['Purch Price of Addtl Purch'].isna(),'Purch Price of Addtl Purch'] = 0

    model_df['Total'] = model_df[['Current Portfolio',
                                      'Addtl Purchase Amt',
                                      'Potential Trades']].sum(axis=1)
    model_df['Now'] = model_df[['Current Portfolio',
                                      'Addtl Purchase Amt']].sum(axis=1)
    
    #create the Blended Actual Purchase Price field
    model_df = BAPP(model_df)

    #create the Blended Price field
    model_df = blended_price(model_df)

    # create the Par Build and Loss fields
    model_df = par_burn_new(model_df)
    
    model_df['Par_no_default'] = model_df['Total'].values
    model_df.loc[model_df['Default']=='Y','Par_no_default'] = 0

    # not using atm
    #model_df['Par_no_default_now'] = model_df['Now'].values
    #model_df.loc[model_df['Default']=='Y','Par_no_default_now'] = 0
    
    return model_df
################################################################

#@xl_func("str: dataframe<index=True>",auto_resize=True)
@xl_func("str: object")
def create_model_port_df(filepath):
    
    # first read in all relevant tables from the CLO model sprdsht
    master_df = get_master_df(filepath,sheet='MASTER')
    CLO_df = get_CLO_df(filepath,sheet=CLO_tab)
    bidask_df = get_bidask_df(filepath,sheet=Bid_tab)
    moodys_score, moodys_rfTable = get_moodys_rating2rf_tables(filepath,sheet='New WARF')
    new_sp_rr, lien_rr, bond_table = get_recovery_rate_tables(filepath,sheet='SP RR Updated')
    #ind_avg_eu = get_ind_avg_eu_table(filepath,sheet='Diversity')
    pot_trades = get_pot_trades(filepath,sheet='Model Portfolio')
    
      
    # merge MASTER + CLO + Bid.Ask + Potential Trades
    model_port = master_df.merge(CLO_df,left_on="LXID",right_on="LXID",how='outer') 
    model_port = model_port.merge(bidask_df,left_on="LXID",right_on="LXID",how='left')
    model_port = model_port.merge(pot_trades, left_on='LXID', right_on='LXID',how='outer')
    model_port.columns = model_port.columns.str.replace('\'','')
    
    # Here is where I think I should add subsetting for needed fields & renaming
    to_rename = {'Potential Trades\nBuys as Positives\nSales as (Negative)':'Potential Trades',
                 'Floating Spread':'Spread',
                 'Floating Spread Floor':'Floor'}
    
    model_port.rename(columns=to_rename,inplace=True)
    
    # add in metric like New WARF, SP's RR, Par no Default, Adj All in Rate
    model_port = moodys_adjusted_warf(model_port,moodys_score,moodys_rfTable)
    model_port = sp_recovery_rate(model_port,lien_rr,new_sp_rr,bond_table)
    
    # need a way to pass LIBOR
    libor = 0.002
    model_port['Adj. All in Rate'] = model_port[['Spread','Floor']].\
        apply(lambda x: (x[0]+x[1]-libor) if (x[1]>libor) else x[0],axis=1 )
    
    #add in all the Pricing stats
    model_port = model_pricing(model_port)
    
    # Let's add dummies for Potential Trades, Current Port for convenience
    #model_port['Pot_Ind'] = abs(model_port['Potential Trades']) > 0
    #model_port['Curr_Ind'] = model_port['Current Portfolio'] > 0
    
    # nil out any tiny positions as they typically don't have full data elsewhere
    # and cause errors, e.g. LX189862 which has 0.03
    model_port = model_port.loc[~((model_port['Current Portfolio']>0)&(model_port['Current Portfolio']<1000))]
      
    #drop the REP Lines
    model_port = model_port.drop(model_port[model_port['Issuer'].str.match('zz_LXREP')].index)
    
    # new, create the marginal stats in main call, same with inside and outside
    model_port = create_marginal_stats(model_port)
    model_port = inside_outside(model_port,['Current', 'Potential'])
        
    return model_port
##################################################################################
@xl_func("dataframe<index=True>,str[] array, float[] array, float[] array: object")
def desirability(model_df,keyStats,weights,highLow):
    def desire(model_df,keyStats,weights,highLow):
            return (((model_df[keyStats]-model_df[keyStats].mean())/model_df[keyStats].std())*\
                    (np.array(weights)*np.array(highLow))).sum(axis=1)
    model_df['Desirability'] = desire(model_df,keyStats,weights,highLow)   
    return model_df
##################################################################################
@xl_func
def inside_outside(model_df,choicelist):
    condlist = [abs(model_df['Current Portfolio']) > 0,            
            abs(model_df['Potential Trades']) > 0]

    choicelist = ['Current', 'Potential']
    model_df['Categorical'] = np.select(condlist, choicelist, default='Outside')
    return model_df

##################################################################################
############ Working Prototype Optimizing Problems ###############################
##################################################################################

@xl_func("dataframe<index=True>, float, float: object", auto_resize=True)
def raise_WAS(model_df,WAS_lift=0.0005,Parburn_crit=2e6):
    """
    model_df: dataframe object
    WAS_lift: Target amount of WAS lift in decimal (5bp is 0.0005)
    Parburn_crit: limit on how much Par to burn
    """
    trade_size = 1e6  # can pass later or make dynamic
    
    if model_df.columns.str.match('Trade').any():
        model_df.drop(columns='Trade',inplace=True)  # need to clear any previous queries
    print("DF: ",type(model_df)," WAS: ",type(WAS_lift)," Parburn: ",type(Parburn_crit))
    ####### Starting Key Stats ####################
    # Let's keep track of pre-trade key stats to compare
    # These should match the pstats from 'Current Portfolio'
    pre_WAS = weighted_average(model_df,cols=['Par_no_default','Spread'])
    target_WAS = pre_WAS + WAS_lift
    pre_WARF = weighted_average(model_df,cols=['Par_no_default','Adj. WARF NEW'])
    pre_dScore = diversity_score(model_df,'Par_no_default')
    print('Pre Trade Stats, WAS: ',pre_WAS,' Target WAS: ',target_WAS,' WARF: ', pre_WARF,' DivScore: ', pre_dScore)

    ####### Potential Trade list ####################
    # technically we just want to rank by highest WAS/lowest parburn  (value/weight)
    pot_sales = model_df.loc[(model_df['Categorical']=='Current')&\
                            (model_df['Spread']<pre_WAS)&(model_df['Adj. WARF NEW']>3200)]\
                            .sort_values(by=['Desirability']).index.tolist() 
    # buys don't have to be outside!
    pot_buys = model_df.loc[(model_df['Categorical']=='Outside')&\
                            (model_df['Spread']>pre_WAS)&(model_df['Adj. WARF NEW']<3200)]\
                            .sort_values(by=['Desirability'],ascending=False).index.tolist()

    num_sales = len(pot_sales)
    num_buys = len(pot_buys)
    num_trades = min(num_sales, num_buys)
    sales_value = model_df.loc[(model_df['Categorical']=='Current')&\
                            (model_df['Spread']<pre_WAS)&(model_df['Adj. WARF NEW']>3200),'Current Portfolio'].sum()
    buys_value = num_buys*trade_size
    #print('# Sales: ',num_sales,' $amt: ', sales_value, ' # Buys: ', num_buys, " $amt: ", buys_value)

    ####### Initial Conditions ####################
    curr_port = model_df.loc[(model_df['Categorical']=='Current')].index.tolist()
    new_port = curr_port.copy()  # start with the current portfolio
    parburn = 0 
    trades = 0
    model_df['PnD_postTrade'] = model_df['Par_no_default']

    trade_bal = 0
    
    for trades in range(min(len(pot_sales),len(pot_buys))):
    
        #let's start with one sale, one buy for now
        # need to actually calculate the burn for actual trade size
#        trade_size = min(model_df.loc[pot_sales[trades],'Par_no_default'],1e6)
        sell_size = min(model_df.loc[pot_sales[trades],'Par_no_default'],trade_size)
        buy_size = sell_size  #this was trade_size
        
        #use this to track how far off cash neutral we are
        trade_bal += buy_size - sell_size
        
        parburn_sale = model_df.loc[pot_sales[trades],'Mil_Par_BL_Sale']
        if parburn_sale > 0:
            buy_size += parburn_sale
        elif parburn_sale < 0:
            buy_size += parburn_sale

        parburn_buy = model_df.loc[pot_buys[trades],'Mil_Par_BL_Buy']*buy_size/trade_size
        if parburn_buy > 0:
            buy_size += parburn_buy
        elif parburn_buy < 0:
            buy_size += parburn_buy
       
        
        parburn_trades = parburn_sale + parburn_buy
        #print("Trade Set #",trades+1)
        parburn += parburn_sale
        #print(pot_sales[trades],parburn)
        # can't use this because the sales isn't sized the same, need to calc actual parburn
        parburn += parburn_buy
        #print(pot_buys[trades],parburn)
        
        if (parburn <= -Parburn_crit):
            #print(parburn, post_WAS)
            break
        else:
            # need to make cash neutral trades
            # could instead keep a running sum of trades and match buys at the end
            model_df.loc[pot_sales[trades],'PnD_postTrade'] = \
                model_df.loc[pot_sales[trades],'Par_no_default'] - sell_size 
            model_df.loc[pot_buys[trades],'PnD_postTrade'] = \
                model_df.loc[pot_buys[trades],'Par_no_default'] + buy_size  #match the sell amount
            
            model_df.loc[pot_sales[trades],'Trade'] = 'Sale'
            new_port.remove(pot_sales[trades])   
            model_df.loc[pot_buys[trades],'Trade'] = 'Buy'
            new_port.append(pot_buys[trades])   
                                    
            #  THIS IS NEW AND UNTESTED
            #  a new counter each time this is triggered that starts at
            #  the beginning of the list (0)
            if (abs(trade_bal) >= trade_size)&(trade_bal<0):
                # do an extra buy trade to balance it closer
                # can test to make sure liquid enough, if not go to next
                model_df.loc[pot_buys[buys],'PnD_postTrade'] += buy_size
                buys += 1
                #    model_df.loc[pot_buys[buys],'Par_no_default'] + buy_size  #match the sell amount
                #model_df.loc[pot_buys[buys],'Trade'] = 'Buy'
                #new_port.append(pot_buys[buys])
            elif (trade_bal >= trade_size)&(trade_bal<0):
                # do an extra sale trade to balance it close
                # can test to make sure liquid enough, if not go to next
                # need to check whether we can do another sale of previous, often not
                while (model_df.loc[pot_sales[sells],'PnD_postTrade'] <1)&(sells <= len(pot_sales)-1):
                    sells +=1
                model_df.loc[pot_sales[sells],'PnD_postTrade'] -= sell_size
                sells += 1
                #    model_df.loc[pot_sales[trades],'Par_no_default'] - sell_size 
                #model_df.loc[pot_sales[trades],'Trade'] = 'Sale'
                #new_port.remove(pot_sales[trades]) 
            #print("Trade Balance: ",trade_bal)

            # on the last trade we need to balance out the trade_bal
            if (trade_bal < 0): #(trades == num_trades -1) & 
                print('Trying to buy more ',pot_buys[buys], ' Buys #',buys)
                model_df.loc[pot_buys[buys],'PnD_postTrade'] += trade_bal
                trade_bal += trade_bal
                buys += 1
            elif (trade_bal > 0):
                #print('Trying to sell more ',pot_sales[sells], ' Sells #',sells)
                while (model_df.loc[pot_sales[sells],'PnD_postTrade'] <1)&(sells <= len(pot_sales)-1):
                    sells +=1
                model_df.loc[pot_sales[sells],'PnD_postTrade'] -= \
                    min(trade_bal,model_df.loc[pot_sales[sells],'PnD_postTrade'])
                trade_bal -= min(trade_bal,model_df.loc[pot_sales[sells],'PnD_postTrade'])
                sells += 1

            post_WAS = weighted_average(model_df.loc[new_port],cols=['PnD_postTrade','Spread'])
            post_dScore = diversity_score(model_df,'PnD_postTrade')
            post_WARF = weighted_average(model_df,cols=['PnD_postTrade','Adj. WARF NEW'])
            
            #print('Sale: ',pot_sales[trades],' Buy: ', pot_buys[trades])
            #print('Post Stats, WAS: ',post_WAS,' WARF: ', post_WARF,' DivScore: ', post_dScore)
            trades += 1
            
            if (post_WAS >= target_WAS)|(post_dScore <= 50):
                #print('Target met: ',(post_WAS >= target_WAS),' or Diversity breached: ', (post_dScore <= 50))
                #print('Total Par Burn: ',parburn,'', post_WAS, post_dScore)
                break
            
        
    return model_df #new_port #, curr_port #, sales, buys

##############################################################################################
@xl_func("dataframe<index=True>, float, float: object", auto_resize=True)
def lower_WARF(model_df,target_WARF,Parburn_crit=0):
    
    trade_size = 1e6  # can pass later or make dynamic
    
    if model_df.columns.str.match('Trade').any():
        model_df.drop(columns='Trade',inplace=True)  # need to clear any previous queries
    
    ####### Starting Key Stats ####################
    # Let's keep track of pre-trade key stats to compare
    # These should match the pstats from 'Current Portfolio'
    pre_WAS = weighted_average(model_df,cols=['Par_no_default','Spread'])
    
    pre_WARF = weighted_average(model_df,cols=['Par_no_default','Adj. WARF NEW'])

    pre_dScore = diversity_score(model_df,'Par_no_default')
    print('Pre Trade Stats, WARF: ',pre_WARF,' Target WARF: ',target_WARF,' WAS: ', pre_WAS,' DivScore: ', pre_dScore)

    ####### Potential Trade list ####################
    # technically we just want to rank by highest WAS/lowest parburn  (value/weight)
    # however ranking by desirability is much more real world
    sell_screen = (model_df['Categorical']=='Current')&\
                            (model_df['MC WARF']>=0)&(model_df['Mil_Par_BL_Sale']>=0)&(model_df['MC WAS']>=0)
    buy_screen = (model_df['Categorical']=='Outside')&(model_df['MC WARF']<=0)&\
                    (model_df['Mil_Par_BL_Buy']>=0)&(model_df['MC WAS']>=0)
    
    pot_sales = model_df.loc[sell_screen]\
                            .sort_values(by=['Desirability']).index.tolist() 
    # buys don't have to be outside!
    pot_buys = model_df.loc[buy_screen]\
                            .sort_values(by=['Desirability'],ascending=False).index.tolist()
    num_sales = len(pot_sales)
    num_buys = len(pot_buys)
    num_trades = min(num_sales, num_buys)
    sales_value = model_df.loc[sell_screen,'Current Portfolio'].sum()
    buys_value = num_buys*trade_size
    print('# Sales: ',num_sales,' $amt: ', sales_value, ' # Buys: ', num_buys, " $amt: ", buys_value)

    ####### Initial Conditions ####################
    curr_port = model_df.loc[(model_df['Categorical']=='Current')].index.tolist()
    new_port = curr_port.copy()  # start with the current portfolio
    parburn = 0 
    trades = 0
    model_df['PnD_postTrade'] = model_df['Par_no_default']   # starts over, key if doing a second query
    
    trade_bal = 0
    mc_parburn = 0

    for trades in range(num_trades):
    
        #let's start with one sale, one buy for now
        # need to actually calculate the burn for actual trade size
        # print(model_df.loc[pot_sales[trades],'Par_no_default'])
        sell_size = min(model_df.loc[pot_sales[trades],'Par_no_default'],trade_size)
        buy_size = trade_size
        
        #use this to track how far off cash neutral we are
        trade_bal += buy_size - sell_size
        
        parburn_sale = model_df.loc[pot_sales[trades],'Mil_Par_BL_Sale']
        if parburn_sale > 0:
            buy_size += parburn_sale
        elif parburn_sale < 0:
            buy_size -= parburn_sale

        parburn_buy = model_df.loc[pot_buys[trades],'Mil_Par_BL_Buy']*buy_size/trade_size
        if parburn_buy > 0:
            buy_size += parburn_buy
        elif parburn_buy < 0:
            buy_buy -= parburn_buy
       
        
        parburn_trades = parburn_sale + parburn_buy
        print("Trade Set #",trades+1)
        parburn += parburn_sale
        print(pot_sales[trades],parburn)
        # can't use this because the sales isn't sized the same, need to calc actual parburn
        parburn += parburn_buy
        print(pot_buys[trades],parburn)
        
        
        
        sells = buys = 0
    
        if (parburn <= -Parburn_crit):
            print(parburn, post_WAS)
            break
        else:
            # need to make cash neutral trades
            # could instead keep a running sum of trades and match buys at the end
            
            model_df.loc[pot_sales[trades],'PnD_postTrade'] = \
                model_df.loc[pot_sales[trades],'Par_no_default'] - sell_size 
            model_df.loc[pot_buys[trades],'PnD_postTrade'] = \
                model_df.loc[pot_buys[trades],'Par_no_default'] + buy_size #match the sell amount
                          
            model_df.loc[pot_sales[trades],'Trade'] = 'Sale'
            new_port.remove(pot_sales[trades])   
            model_df.loc[pot_buys[trades],'Trade'] = 'Buy'
            new_port.append(pot_buys[trades])
            
            #  THIS IS NEW AND UNTESTED
            #  a new counter each time this is triggered that starts at
            #  the beginning of the list (0)
            if (abs(trade_bal) >= trade_size)&(trade_bal<0):
                # do an extra buy trade to balance it closer
                # can test to make sure liquid enough, if not go to next
                model_df.loc[pot_buys[buys],'PnD_postTrade'] += buy_size
                buys += 1
                #    model_df.loc[pot_buys[buys],'Par_no_default'] + buy_size  #match the sell amount
                #model_df.loc[pot_buys[buys],'Trade'] = 'Buy'
                #new_port.append(pot_buys[buys])
            elif (trade_bal >= trade_size)&(trade_bal<0):
                # do an extra sale trade to balance it close
                # can test to make sure liquid enough, if not go to next
                # need to check whether we can do another sale of previous, often not
                model_df.loc[pot_sales[sells],'PnD_postTrade'] -= sell_size
                sells += 1
                #    model_df.loc[pot_sales[trades],'Par_no_default'] - sell_size 
                #model_df.loc[pot_sales[trades],'Trade'] = 'Sale'
                #new_port.remove(pot_sales[trades]) 
            print("Trade Balance: ",trade_bal)

            # on the last trade we need to balance out the trade_bal
            if (trades == num_trades -1) & (trade_bal < 0):
                model_df.loc[pot_buys[buys],'PnD_postTrade'] += trade_bal
                trade_bal += trade_bal
                buys += 1
            elif (trades == num_trades -1) & (trade_bal > 0):
                print('Trying to sell more ',pot_sales[sells], ' Sells #',sells)
                while (model_df.loc[pot_sales[sells],'PnD_postTrade'] <1)&(sells <= len(pot_sales)-1):
                    sells +=1
                model_df.loc[pot_sales[sells],'PnD_postTrade'] -= \
                    min(trade_bal,model_df.loc[pot_sales[sells],'PnD_postTrade'])
                trade_bal -= min(trade_bal,model_df.loc[pot_sales[sells],'PnD_postTrade'])
                sells += 1
                                    
            post_WAS = weighted_average(model_df.loc[new_port],cols=['PnD_postTrade','Spread'])
            post_dScore = diversity_score(model_df,'PnD_postTrade')
            post_WARF = weighted_average(model_df,cols=['PnD_postTrade','Adj. WARF NEW'])
            
            print('Sale: ',pot_sales[trades],' Buy: ', pot_buys[trades])
            print('Post Stats, WAS: ',post_WAS,' WARF: ', post_WARF,' DivScore: ', post_dScore, 'Parburn: ',parburn)
            trades += 1
            
            if (post_WARF <= target_WARF)|(post_dScore <= 50):
                print('Target met: ',(post_WARF <= target_WARF),' or Diversity breached: ', (post_dScore <= 50))
                print('Total Par Burn: ',parburn,'', post_WARF, post_WAS, post_dScore)
                break
        
    #if trade_bal > 0:
    #    model_df.loc[pot_sales[trades],'PnD_postTrade'] =
    #elif trade_bal < 0:
                
        
    return model_df #new_port, curr_port #, sales, buys
###########################################################################################
@xl_func("dataframe<index=True>: dataframe<index=True>", auto_resize=True)
def show_tradelist(model_df):
    return (model_df.loc[~model_df['Trade'].isna(),['Parent Company','Trade','Par_no_default',
                                             'PnD_postTrade','Spread','Adj. WARF NEW','S&P Recovery Rate (AAA)',
                                             'Desirability']]).sort_values(by=['Trade','Desirability'])

############################################################################################
@xl_func
def liquidity_metrics(model_df):
    """
    Poll the traders, Use daily high/low, etc
    """
    model_df['Liquidity_Sale'] = 1
    model_df['Liquidity_Buy'] = 1
    return model_df
############################################################################################
@xl_func("dataframe<index=True>: dataframe<index=True>",auto_resize=True)
def df_describe(df):
    return df.describe()
############################################################################################
@xl_func #("dataframe<index=True>: ",auto_resize=True)
def df_type(df):
    print("Type: ",type(df))
    return type(df)