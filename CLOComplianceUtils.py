import numpy as np  # probably don't need to load
import pandas as pd

#################################################################################
## hard-coded variables for now, maps and such
#################################################################################

libor = 0.0020

dict_2020_20 = {'Class A-1 Notes':240e6,'Class A-2 Notes':16e6,'Class B-1 Notes':39e6,'Class B-2 Notes':9e6,
               'Class C Notes':24e6,'Class D-1 Notes':16e6,'Class D-2 Notes':8e6,'Class E Notes':12e6}

clo_list = ['CLO 4', 'CLO 5', 'CLO 6',
       'CLO 7', 'CLO 8R', 'CLO 9', 'CLO 10', 'CLO 11',
       'CLO 12', 'CLO 13', 'CLO 14', 'CLO 15',
       'CLO 16', 'CLO 17', 'CLO 18', 'CLO 19',
       'CLO 20', 'CLO 21']  # 'Euro 1', 'Euro 2', 'Euro 3', 'Euro 4'

path = 'Z:/Shared/Risk Management and Investment Technology/CLO Optimization/CLO Reports/'
file = 'Master Position Report.xlsx'
filepath = path + file

pathT = 'Z:/Shared/Risk Management and Investment Technology/CLO Optimization/'
fileT = 'US CLO Triggers as of 6.24.21.xlsx'
filepathT = pathT + fileT

qstats_list = ['Minimum Weighted Average S&P Recovery Rate Test - Class A-1',
    'Minimum Floating Spread Test',
    'Minimum Weighted Average Coupon Test',
    'Moody\'s Diversity Test',
    'Maximum Moody\'s Rating Factor Test',
    'Minimum Weighted Average Moody’s Recovery Rate Test',
    'Weighted Average Life']

coverage_stats = ['Interest Coverage Test - Class A',
       'Interest Coverage Test - Class B',
       'Interest Coverage Test - Class A/B',
       'Interest Coverage Test - Class C',
       'Interest Coverage Test - Class D',
       'Interest Diversion Test', 
       'Overcollateralization Ratio Test - Class A',
       'Overcollateralization Ratio Test - Class B',
       'Overcollateralization Ratio Test - Class A/B',
       'Overcollateralization Ratio Test - Class C',
       'Overcollateralization Ratio Test - Class D',
       'Overcollateralization Ratio Test - Class E',
       'Reinvestment Overcollateralization Test',
       'Overcollateralization Ratio Test - Event of Default']
       
coverage_map = {'Interest Coverage Test - Class A':'Interest Coverage Test - Class A',
       'Interest Coverage Test - Class B':'Interest Coverage Test - Class B',
       'Interest Coverage Test - Class A/B':'Interest Coverage Test - Class A/B',
       'Interest Coverage Test - Class C':'Interest Coverage Test - Class C',
       'Interest Coverage Test - Class D':'Interest Coverage Test - Class D',
       'Interest Diversion Test':'Interest Diversion Test', 
       'Overcollateralization Ratio Test - Class A':'Overcollateralization Ratio Test - Class A',
       'Overcollateralization Ratio Test - Class B':'Overcollateralization Ratio Test - Class B',
       'Overcollateralization Ratio Test - Class A/B':'Overcollateralization Ratio Test - Class A/B',
       'Overcollateralization Ratio Test - Class C':'Overcollateralization Ratio Test - Class C',
       'Overcollateralization Ratio Test - Class D':'Overcollateralization Ratio Test - Class D',
       'Overcollateralization Ratio Test - Class E':'Overcollateralization Ratio Test - Class E',
       'Reinvestment Overcollateralization Test':'Reinvestment Overcollateralization Test',
       'Overcollateralization Ratio Test - Class A1 [Event of Default - Section 5.1(g)':
                'Overcollateralization Ratio Test - Event of Default',
       'Overcollateralization Ratio Test - Event of Default':'Overcollateralization Ratio Test - Event of Default',
       'Class B Interest Coverage Test':'Interest Coverage Test - Class B',  #
       'Class C Interest Coverage Test': 'Interest Coverage Test - Class C',
       'Class D Interest Coverage Test':'Interest Coverage Test - Class D',
       'Class A Interest Coverage Test':'Interest Coverage Test - Class A',
       'Class B Overcollateralization Ratio Test':'Overcollateralization Ratio Test - Class B',
       'Class C Overcollateralization Ratio Test':'Overcollateralization Ratio Test - Class C',
       'Class D Overcollateralization Ratio Test':'Overcollateralization Ratio Test - Class D',
       'Class A Overcollateralization Ratio Test':'Overcollateralization Ratio Test - Class A',
       'Class E Overcollateralization Ratio Test':'Overcollateralization Ratio Test - Class E',
       'Class A-1 Overcollateralization Ratio Test [Event of Default - Section 5.1(g)]':
               'Overcollateralization Ratio Test - Event of Default'}

con_stats = ['Cov-Lite Loans','Moody\'s Default Probability Rating <= Caa1 and/or S&P Rating <= CCC+',
    'Moody\'s Rating <= Caa1','S&P Rating <= CCC+']

con_stats_map = {'Cov-Lite Loans':'Cov-Lite Loans','Moody\'s Default Probability Rating Of Caa1 or Below':
                 'Moody\'s Rating <= Caa1','Moody\'s Default Probability Rating Of Caa1 or Below and/or S&P Rating of CCC+ or Below':
                 'Moody\'s Default Probability Rating <= Caa1 and/or S&P Rating <= CCC+',
                 'Moody\'s Rating <= Caa1':'Moody\'s Rating <= Caa1','S&P Rating <= CCC+':'S&P Rating <= CCC+',
                 'S&P Rating of CCC+ or Below':'S&P Rating <= CCC+'}

all_stats_map = {'Min Floating Spread Test - no Libor Floors':'Minimum Floating Spread Test',
       'Min Floating Spread Test - With Libor Floors':'Min Floating Spread Test - With Libor Floors',
       'Minimum Weighted Average Coupon Test':'Minimum Weighted Average Coupon Test',
       'Max Moodys Rating Factor Test (NEW WARF)':'Max Moodys Rating Factor Test (NEW WARF)',
       'Max Moodys Rating Factor Test (Orig WARF)':'Maximum Moody\'s Rating Factor Test',
       'Min Moodys Recovery Rate Test':'Minimum Weighted Average Moody’s Recovery Rate Test', 
       'Min S&P Recovery Rate Class A-1a':'Minimum Weighted Average S&P Recovery Rate Test - Class A-1',
       'Moodys Diversity Test':'Moody\'s Diversity Test', 
       'Weighted Average Life Test':'Weighted Average Life',
       'Percent Caa & lower':'Moody\'s Rating <= Caa1', 
       'Percent CCC & lower':'S&P Rating <= CCC+', 
       'Percent 2nd Lien':'Percent 2nd Lien',
       'Total Portfolio Par (excl. Defaults)':'Total Portfolio Par (excl. Defaults)',
       'Class A/B Overcollateralization Ratio':'Overcollateralization Ratio Test - Class A/B',
       'Class C Overcollateralization Ratio':'Overcollateralization Ratio Test - Class C',
       'Class D Overcollateralization Ratio':'Overcollateralization Ratio Test - Class D',
       'Class E Overcollateralization Ratio':'Overcollateralization Ratio Test - Class E',
       'Reinvestment Overcollateralization Ratio':'Reinvestment Overcollateralization Test',
       'S&P Weighted Average Rating Factor (SP WARF)':'S&P Weighted Average Rating Factor (SP WARF)',
       'Default Rate Dispersion (DRD)':'Default Rate Dispersion (DRD)', 
       'Obligor Diversity Measure (ODM)':'Obligor Diversity Measure (ODM)',
       'Industry Diversity Measure (IDM)':'Industry Diversity Measure (IDM)', 
       'Regional Diversity Measure (RDM)':'Regional Diversity Measure (RDM)',
       'Weighted Average Life (WAL)':'Weighted Average Life (WAL)', 
       'Break-Even Default Rate (BDR)':'Break-Even Default Rate (BDR)',
       'Adjusted Break-Even Default Rate (Adj BDR)':'Adjusted Break-Even Default Rate (Adj BDR)',
       'Scenario Default Rate (SDR)':'Scenario Default Rate (SDR)'}


#################################################################################
## These functions have been modified vis a vis CLOUtils due to differences in
## input files.  The CLOUtils will need to be updated to match these soon but 
## these are "forked" to get going quickly, and the optimizer's version will
## need to work with these versions going forward
#################################################################################
def Port_stats(model_df, weight_col='Par_no_default',format_output=False):
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
        'Minimum Weighted Average Coupon Test',
        'Max Moodys Rating Factor Test (NEW WARF)',
        'Max Moodys Rating Factor Test (Orig WARF)',
        'Min Moodys Recovery Rate Test',
        'Min S&P Recovery Rate Class A-1a',
        'Moodys Diversity Test',
    #    'WAP',
        'Weighted Average Life Test',
        'Percent Caa & lower',
        'Percent CCC & lower',
        'Percent 2nd Lien',
    #    'Percent Sub80',
    #    'Percent Sub90',
        'Percent CovLite',
    #    'Pot Par B/L Sale',
    #    'Pot Par B/L Buy',
    #    'Pot Par B/L Total',
        'Total Portfolio Par (excl. Defaults)',
    #    'Total Portfolio Par',
     #   'Current Portfolio'
                                              ],columns = [weight_col])

    
    # Par_no_default works for current positions only
    # if trying to compare potential new trades, need updated field
    # also Total isn't dynamically updated for Potential Trades once
    # model_df is created.  Need to find solution
    
    Port_stats_df.loc['Min Floating Spread Test - no Libor Floors',weight_col] = \
        Weighted_Average_Spread(model_df,weight_col,libor=.002)*100
    
    Port_stats_df.loc['Min Floating Spread Test - With Libor Floors',weight_col] = \
        weighted_average(model_df,cols=[weight_col,'All In Rate'])*100

    Port_stats_df.loc['Minimum Weighted Average Coupon Test',weight_col] = \
        Weighted_Average_Coupon(model_df,weight_col)*100

    Port_stats_df.loc['Max Moodys Rating Factor Test (NEW WARF)',weight_col] = \
        weighted_average(model_df,cols=[weight_col,'Adj. WARF NEW'])
        #moodys_adjusted_warf(model_df)  #Adj. WARF NEW
     
    Port_stats_df.loc['Max Moodys Rating Factor Test (Orig WARF)',weight_col] = \
        weighted_average(model_df,cols=[weight_col,'WARF'])
        #moodys_adjusted_warf_old(model_df)
    
    
        
    Port_stats_df.loc['Min Moodys Recovery Rate Test',weight_col] = \
        weighted_average(model_df,cols=[weight_col,'Moodys Recovery Rate'])*100    
    
    Port_stats_df.loc['Min S&P Recovery Rate Class A-1a',weight_col] = \
            weighted_average(model_df,cols=[weight_col,'S&P Recovery Rate (AAA)'])*100
        #weighted_average(model_df,cols=[weight_col,'S&P Recovery Rate (AAA)'])*100
    
    Port_stats_df.loc['Moodys Diversity Test',weight_col] = diversity_score(model_df, weight_col)
    
    ######   Was Total?
    #Port_stats_df.loc['WAP',weight_col] = \
    #    weighted_average(model_df,cols=[weight_col,'Blended Price'])
    
    # it needs a different amort table for each CLO, so this is place holder
    Port_stats_df.loc['Weighted Average Life Test',weight_col] = \
            Weighted_Average_Life(model_df,weight_col)
    
    Port_stats_df.loc['Percent Caa & lower',weight_col] = percentage_Caa(model_df, weight_col)*100
    Port_stats_df.loc['Percent CCC & lower',weight_col] = percentage_CCC(model_df, weight_col)*100
    
    Port_stats_df.loc['Percent 2nd Lien',weight_col] = percentage_SecondLien(model_df, weight_col)*100
    
    #Port_stats_df.loc['Percent Sub80',weight_col] = percentage_SubEighty(model_df, weight_col)*100
    
    #Port_stats_df.loc['Percent Sub90',weight_col] = percentage_SubNinety(model_df, weight_col)*100
    
    Port_stats_df.loc['Percent CovLite',weight_col] = percentage_CovLite(model_df, weight_col)*100
    Port_stats_df.loc['Total Portfolio Par (excl. Defaults)',weight_col] = model_df[weight_col].sum()

    
    if format_output:
        Port_stats_df.loc['Min Floating Spread Test - no Libor Floors'] = \
            Port_stats_df.loc['Min Floating Spread Test - no Libor Floors'].apply('{:.2f}%'.format)
        Port_stats_df.loc['Min Floating Spread Test - With Libor Floors'] = \
            Port_stats_df.loc['Min Floating Spread Test - With Libor Floors'].apply('{:.2f}%'.format)
        Port_stats_df.loc['Minimum Weighted Average Coupon Test'] = \
            Port_stats_df.loc['Minimum Weighted Average Coupon Test'].apply('{:.2f}%'.format)
        Port_stats_df.loc['Max Moodys Rating Factor Test (NEW WARF)'] = \
            Port_stats_df.loc['Max Moodys Rating Factor Test (NEW WARF)'].apply('{:.0f}'.format)
        Port_stats_df.loc['Max Moodys Rating Factor Test (Orig WARF)'] = \
            Port_stats_df.loc['Max Moodys Rating Factor Test (Orig WARF)'].apply('{:.0f}'.format)
        Port_stats_df.loc['Min Moodys Recovery Rate Test'] = \
            Port_stats_df.loc['Min Moodys Recovery Rate Test'].apply('{:.1f}%'.format)
        Port_stats_df.loc['Min S&P Recovery Rate Class A-1a'] = \
            Port_stats_df.loc['Min S&P Recovery Rate Class A-1a'].apply('{:.1f}%'.format)
        Port_stats_df.loc['Moodys Diversity Test'] = \
            Port_stats_df.loc['Moodys Diversity Test'].apply('{:.0f}'.format)
        #Port_stats_df.loc['WAP'] = \
        #    Port_stats_df.loc['WAP'].apply('${:.2f}'.format)
        # Weighted_Average_Coupon(clo_df,col)
        Port_stats_df.loc['Weighted Average Life Test'] = \
            Port_stats_df.loc['Weighted Average Life Test'].apply('{:.2f}'.format)
        Port_stats_df.loc['Percent Caa & lower'] = \
            Port_stats_df.loc['Percent Caa & lower'].apply('{:.1f}%'.format)
        Port_stats_df.loc['Percent CCC & lower'] = \
            Port_stats_df.loc['Percent CCC & lower'].apply('{:.1f}%'.format)
        Port_stats_df.loc['Percent 2nd Lien'] = \
            Port_stats_df.loc['Percent 2nd Lien'].apply('{:.1f}%'.format)
        #Port_stats_df.loc['Percent Sub80'] = \
        #    Port_stats_df.loc['Percent Sub80'].apply('{:.1f}%'.format)
        #Port_stats_df.loc['Percent Sub90'] = \
        #    Port_stats_df.loc['Percent Sub90'].apply('{:.1f}%'.format)
        Port_stats_df.loc['Percent CovLite'] = \
            Port_stats_df.loc['Percent CovLite'].apply('{:.1f}%'.format)
        Port_stats_df.loc['Total Portfolio Par (excl. Defaults)'] = \
            Port_stats_df.loc['Total Portfolio Par (excl. Defaults)'].apply('{:,.0f}'.format)
        
  
    #Port_stats_df.loc['Pot Par B/L Sale',weight_col] = model_df['Par_Build_Loss_Sale'].sum()
    #Port_stats_df.loc['Pot Par B/L Sale'] = \
    #    Port_stats_df.loc['Pot Par B/L Sale'].apply('{:,.0f}'.format)    
    
    #Port_stats_df.loc['Pot Par B/L Buy',weight_col] = model_df['Par_Build_Loss_Buy'].sum()
    #Port_stats_df.loc['Pot Par B/L Buy'] = \
    #    Port_stats_df.loc['Pot Par B/L Buy'].apply('{:,.0f}'.format)    
    
    #Port_stats_df.loc['Pot Par B/L Total',weight_col] = model_df['Total_Par_Build_Loss'].sum()
    #Port_stats_df.loc['Pot Par B/L Total'] = \
    #    Port_stats_df.loc['Pot Par B/L Total'].apply('{:,.0f}'.format)    

    
#    Port_stats_df.loc['Total Portfolio Par',weight_col] = model_df['Total'].sum()
#    Port_stats_df.loc['Total Portfolio Par'] = \
#        Port_stats_df.loc['Total Portfolio Par'].apply('{:,.0f}'.format)
    
    # current portfolio is Quantity + Add'l Amount (manual) TBA later
#    Port_stats_df.loc['Current Portfolio',weight_col] = model_df[['Addtl Purchase Amt','Current Portfolio']].sum(axis=1).sum()
#    Port_stats_df.loc['Current Portfolio'] = \
#        Port_stats_df.loc['Current Portfolio'].apply('{:,.0f}'.format)
    
    return Port_stats_df

#################################################################################
def DataLabelMap(df):
    datalabelmap = {'Parent Company': 'Parent Company',
        'Issuer': 'Issuer',
        'Asset': 'Asset',
        'CUSIP': 'CUSIP',
        'ISIN': 'ISIN',
        'Asset Type': 'Asset Type',
        'Analyst': 'Analyst',
        'Floating Spread': 'Spread',
        'Floating Spread Floor': 'Floor',
        'All In Rate': 'All In Rate',
        'Maturity Date': 'Maturity Date',
        'Mark Price': 'Mark Price',
        'Adjusted Moodys Rating': 'Adjusted Moodys Rating',
        'WARF': 'WARF',
        'MDPR': 'MDPR',
        'Moody\'s CFR': 'Moodys CFR',
        'Moody\'s Facility Rating': 'Moodys Facility Rating',
        'Moody\'s Issuer Outlook': 'Moodys Issuer Outlook',
        'Moody\'s Issuer Watch': 'Moodys Issuer Watch',
        'Moodys Recovery Rate': 'Moodys Recovery Rate',
        'S&P Issuer Rating': 'S&P Issuer Rating',
        'S&P Facility Rating': 'S&P Facility Rating',
        'S&P Issuer Outlook': 'S&P Issuer Outlook',
        'S&P Issuer Watch': 'S&P Issuer Watch',
        'S&P Recovery': 'S&P Recovery',
        'Agent Bank': 'Agent Bank',
        'OCP CLO 2013-4, Ltd.': 'CLO 4',
        'OCP CLO 2014-5, Ltd.': 'CLO 5',
        'OCP CLO 2014-6, Ltd.': 'CLO 6',
        'OCP CLO 2014-7, Ltd.': 'CLO 7',
        'OCP CLO 2020-8R, Ltd.': 'CLO 8R',
        'OCP CLO 2015-9, Ltd.': 'CLO 9',
        'OCP CLO 2015-10, Ltd.': 'CLO 10',
        'OCP CLO 2016-11, Ltd.': 'CLO 11',
        'OCP CLO 2016-12, Ltd.': 'CLO 12',
        'OCP CLO 2017-13, Ltd.': 'CLO 13',
        'OCP CLO 2017-14, Ltd.': 'CLO 14',
        'OCP CLO 2018-15, Ltd.': 'CLO 15',
        'OCP CLO 2019-16, Ltd.': 'CLO 16',
        'OCP CLO 2019-17, Ltd': 'CLO 17',
        'OCP CLO 2020-18, Ltd.': 'CLO 18',
        'OCP CLO 2020-19, Ltd.': 'CLO 19',
        'OCP CLO 2020-20, Ltd.': 'CLO 20',
        'OCP CLO 2021-21 Ltd.': 'CLO 21',
        'Current Global Amount Outstanding': 'Current Global Amount Outstanding',
        'Moody\'s Industry': 'Moodys Industry',
        'S&P Industry': 'S&P Industry',
        'Lien Type': 'Lien Type',
        'Issuer Country': 'Issuer Country'}
    df.columns = df.columns.str.strip()  # removes leading/trailing whitespace
    df.rename(columns=datalabelmap,inplace=True)
    return df
#################################################################################
def get_master_position_report(filepath):
    master_df = pd.read_excel(filepath,header=0)
    master_df = master_df.loc[:,~master_df.columns.str.match("Unnamed")]
    master_df.rename(columns={'LoanX ID':'LXID'},inplace=True)
    master_df.set_index('LXID', inplace=True)
    master_df = DataLabelMap(master_df)
    master_df = add_derived_features(master_df,libor=0.0020)
    return master_df
#################################################################################
def diversity_score(model_df, weight_col='Par_no_default'):
    """
    This function calculates the Moody's Industry Diversity Score for the CLO
    
    Arg in:
        model_df: the input data frame (from the MASTER table d/l'd from BMS)
        #ind_avg_eu: Moody's discrete lookup table that maps AIEUS to IDS, need to be sorted
    Arg out:
        dscore: the scalar measure of the IDS
    """
    # calling this each time (like in MC Diversity)
    # makes the function unbearably slow. I tried setting it
    # as a global variable but that didn't solve the issue.
    # I need to look into a solution to set something up 
    # permanently in memory after the first call (thought global would do it)
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
    
    # 6/14/2021 added this...it is correct
    # filters out the zero names from the average Par Amount (in the count)
    mask = abs(model_df[weight_col]) > 0
    
    
    div_df = model_df.loc[mask,['Parent Company','Moodys Industry',weight_col]].copy()
    div_df.sort_values(by='Moodys Industry',inplace=True)

    # this keeps the industry, but groups on parent company for multiple loans
    # shouldn't this be Obligor for Diversity field and not Parent Company?
    # they are the same in our spreadsheet, but maybe need to be careful
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
#################################################################################
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
#################################################################################
def SP_CDO_Coefs(col='CLO 20'):
    """
    Looks up the coefficients in the dict table and returns the 3 as a float array
    """
        
    clo_coefs = {'CLO 18': [0.121609,3.706211,1.025031],
                'CLO 16': [0.101745,4.769688,0.962975],
                'CLO 19': [0.097805,4.15832,1.071661],
                'CLO 17': [0.061931,4.266469,1.104848],
                'CLO 20': [0.109706,4.140929,1.064031],
                'CLO 8R': [0.106769,3.967347,0.997703],
                'CLO 14': [0.093957,4.215527,1.086073],
                'CLO 13': [0.066717,4.243131,1.11317],
                'CLO 15': [0.065108,4.620284,1.048429],
                'CLO 10': [-0.008128,3.480512,1.260735],
                'CLO 11': [0.059116,4.481791,1.10217],
                'CLO 12': [-0.046444,3.833393,1.107043],
                'CLO 7': [0.017043,4.541246,1.067143],
                'CLO 6': [0.05869,4.267761,1.099711],
                'CLO 5': [0.060439,4.275955,1.14419],
                'EUR 1': [0.185893082941552,3.70021989750327,0.881097473356921],
                'EUR 2': [0.206296268943602,4.05917015316102,0.947621203737026],
                'EUR 3': [0.191101158723369,3.39938112043418,0.952667210171681],
                'EUR 4': [np.nan,np.nan,np.nan]}
    
    return clo_coefs[col]
#################################################################################
def SP_CDO_Type(col='CLO 20'):
    """
    There seems to be two types of BDR test stats
    Type II: 0.247621 + (SPWARF/9162.65) – (DRD/16757.2) – (ODM/7677.8) – (IDM/2177.56) – (RDM/34.0948) + (WAL/27.3896)
    Type I: 0.329915 + (1.210322 * EPDR) – (0.586627 * DRD) + (2.538684 /ODM) + (0.216729 / IDM) + (0.0575539 / RDM) – (0.0136662 * WAL)
    """
        
    clo_types = {'CLO 18': 'Type II',
                'CLO 16': 'Type II',
                'CLO 19': 'Type II',
                'CLO 17': 'Type II',
                'CLO 20': 'Type II',
                'CLO 8R': 'Type II',
                'CLO 14': 'Type I',
                'CLO 13': 'Type I',
                'CLO 15': 'Type I',
                'CLO 10': 'Type I',
                'CLO 11': 'Type I',
                'CLO 12': 'Type I',
                'CLO 7': 'Type I',
                'CLO 6': 'Type I',
                'CLO 5': 'Type I',
                'EUR 1': 'Type II',
                'EUR 2': 'Type I',
                'EUR 3': 'Type II',
                'EUR 4': 'Type II'}
    
    return clo_types[col]
#################################################################################
def get_SPDR_Table(filepath):
    filepath = 'Z:/Shared/Risk Management and Investment Technology/CLO Optimization/US CLO Triggers as of 6.24.21.xlsx'
    SPDR_Table = pd.read_excel(filepath,sheet_name = "SPDR")
    SPDR_Table.set_index('Tenor', inplace=True)
    return SPDR_Table
#################################################################################
def SPDR(df,SPDR_Table):
    
    df['Loan Life']  = ((df['Maturity Date'] - pd.Timestamp.today()).dt.days/365)
    
    badrates = pd.Series(['CC','NR'])
    
    def lookup_spdr(rating,loanlife):
        # this needs to skip any sub CCC- loans
        if rating!=rating:  # checking string is nan trick
            spdr = np.nan
        elif badrates.str.match(rating).any():
            spdr = np.nan
        else:
            tenors = (SPDR_Table.columns == loanlife//1) | (SPDR_Table.columns == loanlife//1 +1)
            spdr = SPDR_Table.loc[rating,tenors].values[0]+ loanlife%1 * \
                (SPDR_Table.loc[rating,tenors].values[1]-SPDR_Table.loc[rating,tenors].values[0])
        return df
        
    df['S&P Default Rating'] = df[['S&P Facility Rating','Loan Life']].apply(lambda x: lookup_spdr(x[0],x[1]),axis=1)
    
    return df
#################################################################################
def SP_CDO_Monitor_Test(clo_df,col='CLO 20'):

    C0=SP_CDO_Coefs(col)[0]
    C1=SP_CDO_Coefs(col)[1]
    C2=SP_CDO_Coefs(col)[2]
    
    # 'S&P Issuer Rating'
    #clo_df.drop(columns=['S&P CLO Specified Assets'],inplace=True)
    clo_df['S&P CLO Specified Assets'] = clo_df[col].values
    mask = (clo_df['S&P Issuer Rating']=='CC') | \
           (clo_df['S&P Issuer Rating']=='SD') | \
           (clo_df['S&P Issuer Rating']=='NR') | \
           (clo_df['S&P Issuer Rating']=='D')
    clo_df.loc[mask,'S&P CLO Specified Assets'] = 0   # what about .isna()?
   
    # Target Par Amount
    # this is a quick estimate until we have a default parameter set for each
    OP = (round((clo_df[col].sum()/1e8))*100)*1e6
    
    # Agg Principal Balance (excl < CCC-)
    NP = clo_df['S&P CLO Specified Assets'].sum()
    
    print("OP: ", OP, ", NP: ", NP)
    # Weighted Average Spread
    SPWAS = weighted_average(clo_df,cols=[col,'Spread'])
    
    # Weighted Average Recovery Rate
    # uses par no default basically
    SPWARR = weighted_average(clo_df,cols=[col,'S&P Recovery Rate (AAA)'])
    
    # Weighted Average Rating Factor
    # uses specified assets
    SPWARF = weighted_average(clo_df,cols=['S&P CLO Specified Assets','S&P Global Ratings Factor'])
    
    
    clo_df['Absolute Deviation'] = abs(clo_df['S&P Global Ratings Factor']-SPWARF)
    # Default Rate Dispersion
    # weighted average of |Global Rating Factor - WARF| (excl < CCC-)
    # use specified assets
    DRD = weighted_average(clo_df,cols=['S&P CLO Specified Assets','Absolute Deviation'])
    
    #print("SPWAS: ", SPWAS, ", SPWARR: ", SPWARR, ', SPWARF: ',SPWARF, ", DRD: ", DRD)
    
    # Obligor Diversity Measure
    vec = clo_df[[col,'Parent Company']].groupby('Parent Company').sum().values/NP
    ODM = 1/vec.T.dot(vec)
    
    # Industry Diversity Measure
    vec = clo_df[[col,'S&P Industry']].groupby('S&P Industry').sum().values/NP
    IDM = 1/vec.T.dot(vec)
    
    # Regional Diversity Measure
    # Annex D : S&P CDO Evaluator Country Codes, Regions and Recovery Groups
    vec = clo_df[[col,'Issuer Country']].groupby('Issuer Country').sum().values/NP
    RDM = 1/vec.T.dot(vec)
    
    # Weighted Average Life
    WAL = Weighted_Average_Life(clo_df,col)
    
    #print("ODM: ", ODM, ", IDM: ", IDM, ', RDM: ',RDM, ", WAL: ", WAL)
    
    # S&P CDO Monitor BDR
    BDR = C0 + (C1 * SPWAS) + (C2 * SPWARR)
    
    # S&P CDO Monitor Adjusted BDR
    AdjBDR = BDR * (OP/NP) - (NP - OP)/(NP *(1-SPWARR))
    
    # S&P CDO Monitor SDR
    if SP_CDO_Type(col) == 'Type II':
        SDR = 0.247621 + (SPWARF/9162.65) - (DRD/16757.2) - (ODM/7677.8) - (IDM/2177.56) - (RDM/34.0948) \
            + (WAL/27.3896)
    else:
        SPDR = 
        
        # S&P Expected Portfolio Default Rate
        EPDR = (model_df['S&P CLO Specified Assets']*SPDR).sum()/model_df['S&P CLO Specified Assets'].sum()

        SDR = 0.329915 + (1.210322 * EPDR) – (0.586627 * DRD) + (2.538684 /ODM) + (0.216729 / IDM) + \
            (0.0575539 / RDM) – (0.0136662 * WAL)
    
    print('Pass' if AdjBDR >= SDR else 'Fail')
    
    SP_CDO_mon_df = pd.DataFrame(np.nan,index=['S&P Weighted Average Rating Factor (SP WARF)',
        'Default Rate Dispersion (DRD)',
        'Obligor Diversity Measure (ODM)',
        'Industry Diversity Measure (IDM)',
        'Regional Diversity Measure (RDM)',
        'Weighted Average Life (WAL)',
        'Break-Even Default Rate (BDR)',
        'Adjusted Break-Even Default Rate (Adj BDR)',
        'Scenario Default Rate (SDR)'],columns = [col])

    SP_CDO_mon_df.loc['S&P Weighted Average Rating Factor (SP WARF)',col] =  SPWARF
    SP_CDO_mon_df.loc['Default Rate Dispersion (DRD)',col] = DRD
    SP_CDO_mon_df.loc['Obligor Diversity Measure (ODM)',col] = ODM
    SP_CDO_mon_df.loc['Industry Diversity Measure (IDM)',col] = IDM
    SP_CDO_mon_df.loc['Regional Diversity Measure (RDM)',col] = RDM
    SP_CDO_mon_df.loc['Weighted Average Life (WAL)',col] = WAL
    SP_CDO_mon_df.loc['Break-Even Default Rate (BDR)',col] = BDR
    SP_CDO_mon_df.loc['Adjusted Break-Even Default Rate (Adj BDR)',col] = AdjBDR
    SP_CDO_mon_df.loc['Scenario Default Rate (SDR)',col] = SDR
    
    return SP_CDO_mon_df
#################################################################################
def moodys_adjusted_warf(df):
    """
    This function creates the new Moody's Ratings Factor based 
    on the old Moody's rating.
    
    Arg in:
        df: the input data frame (from the MASTER table d/l'd from BMS)     
    """
    
    # moodys_score: dataframe with alphanumeric rating to numeric map (1 to 1 map; linear)
    moodys_score = pd.DataFrame([[ 'Aaa',1],['Aa1',2],['Aa2',3],['Aa3',4],
              ['A1',5],['A2',6],['A3',7],['Baa1',8],['Baa2',9],
              ['Baa3',10],['Ba1',11],['Ba2',12],['Ba3',13],
              ['B1',14],['B2',15],['B3',16],['Caa1',17],
              ['Caa2',18],['Caa3',19],['Ca',20],['C',21]],columns=['Moodys','Score'])
    
    # moodys_rfTable: dataframe with alphanumeric rating to new WARF numeric (1 to 1 map; 1 to 1000 values)
    moodys_rfTable = pd.DataFrame([[ 'Aaa',1],['Aa1',10],['Aa2',20],['Aa3',40],
              ['A1',70],['A2',120],['A3',180],['Baa1',260],['Baa2',360],
              ['Baa3',610],['Ba1',940],['Ba2',1350],['Ba3',1766],
              ['B1',2220],['B2',2720],['B3',3490],['Caa1',4770],
              ['Caa2',6500],['Caa3',8070],['Ca',10000],['C',10000],
              ['NR',10000],['W/D',10000]],columns=['Moodys Rating Factor Table','RF'])
    
    score = df['Moodys CFR'].map(dict(moodys_score[['Moodys','Score']].values))
    updown = df['Moodys Issuer Watch'].\
        apply(lambda x: -1 if x == 'Possible Upgrade' else 1 if x == 'Possible Downgrade' else 0)
    aScore = score + updown
    Adjusted_CFR_for_WARF = aScore.map(dict(moodys_score[['Score','Moodys']].values))
    # I keep the same column name as Jeff to make it easier to double check values
    df['Adj. WARF NEW'] = Adjusted_CFR_for_WARF.map(dict(moodys_rfTable[['Moodys Rating Factor Table','RF']].values))
    return df

#################################################################################
def moodys_adjusted_warf_old(df):
    """
    This function creates the new Moody's Ratings Factor based 
    on the old Moody's rating. 
    
    !!! This needs to have the old WARF logic added, just a place holder ATM !!!
    
    Arg in:
        df: the input data frame (from the MASTER table d/l'd from BMS)     
    """
    
    # moodys_score: dataframe with alphanumeric rating to numeric map (1 to 1 map; linear)
    moodys_score = pd.DataFrame([[ 'Aaa',1],['Aa1',2],['Aa2',3],['Aa3',4],
              ['A1',5],['A2',6],['A3',7],['Baa1',8],['Baa2',9],
              ['Baa3',10],['Ba1',11],['Ba2',12],['Ba3',13],
              ['B1',14],['B2',15],['B3',16],['Caa1',17],
              ['Caa2',18],['Caa3',19],['Ca',20],['C',21]],columns=['Moodys','Score'])
    
    # moodys_rfTable: dataframe with alphanumeric rating to new WARF numeric (1 to 1 map; 1 to 1000 values)
    moodys_rfTable = pd.DataFrame([[ 'Aaa',1],['Aa1',10],['Aa2',20],['Aa3',40],
              ['A1',70],['A2',120],['A3',180],['Baa1',260],['Baa2',360],
              ['Baa3',610],['Ba1',940],['Ba2',1350],['Ba3',1766],
              ['B1',2220],['B2',2720],['B3',3490],['Caa1',4770],
              ['Caa2',6500],['Caa3',8070],['Ca',10000],['C',10000],
              ['NR',10000],['W/D',10000]],columns=['Moodys Rating Factor Table','RF'])
    
    score = df['Moodys CFR'].map(dict(moodys_score[['Moodys','Score']].values))
    updown = df['Moodys Issuer Watch'].\
        apply(lambda x: -1 if x == 'Possible Upgrade' else 1 if x == 'Possible Downgrade' else 0)
    aScore = score + updown
    Adjusted_CFR_for_WARF = aScore.map(dict(moodys_score[['Score','Moodys']].values))
    # I keep the same column name as Jeff to make it easier to double check values
    df['Adj. WARF NEW'] = Adjusted_CFR_for_WARF.map(dict(moodys_rfTable[['Moodys Rating Factor Table','RF']].values))
    return df

#################################################################################
def sp_grf(df):
    sp_grf = pd.DataFrame([['AAA',13.51],['AA+',26.75],['AA',46.36],['AA-',63.90],
                            ['A+',99.50],['A',146.35],['A-',199.83],['BBB+',271.01],
                            ['BBB',361.17],['BBB-',540.42],['BB+',784.92],['BB',1233.63],
                            ['BB-',1565.44],['B+',1982.00],['B',2859.50],['B-',3610.11],
                            ['CCC+',4641.40],['CCC',5293.00],['CCC-',5751.00],['CC',10000.00],
                            ['SD',10000.00],['D',10000.00]],columns = ['S&P Issuer Rating','S&P Global Ratings Factor'])
    df['S&P Global Ratings Factor'] = df['S&P Issuer Rating'].map(dict(sp_grf[['S&P Issuer Rating','S&P Global Ratings Factor']].values))
    return df
#################################################################################
def add_derived_features(df,libor=0.0020):
    df = moodys_adjusted_warf(df)   # 'Adj. WARF NEW'
    df = sp_recovery_rate(df)       # 'S&P Recovery Rate (AAA)'
    df = sp_grf(df)                 # 'S&P Global Ratings Factor'
    df = Effective_Spread(df,libor=0.0020)  # Calc the EffSpread = Spread + max(floor-libor,0)

    return df
#################################################################################
def sp_recovery_rate(model_df):
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
    new_rr_map = {'1+(100)': 0.75,
                '1(95%)': 0.70,
                '1(90%)': 0.65,
                '2(85%)': 0.625,
                '2(80%)': 0.60,
                '2(75%)': 0.55,
                '2(70%)': 0.5,
                '3(65%)': 0.45,
                '3(60%)': 0.4,
                '3(55%)': 0.35,
                '3(50%)': 0.3,
                '4(45%)': 0.285,
                '4(40%)': 0.27,
                '4(35%)': 0.235,
                '4(30%)': 0.20,
                '5(25%)': 0.175,
                '5(20%)': 0.15,
                '5(15%)': 0.10,
                '5(10%)': 0.05,
                '6(5%)': 0.035,
                '6(0%)': 0.02,
                '3H': 0.40,
                '1': 0.65}
    
    LienOne_map = {'AU':0.50,'AT':0.50,'BE':0.50,
            'CA':0.50,'DK':0.50,'FI':0.50,'FR':0.50,
            'DE':0.50,'HK':0.50,'IE':0.50,'IS':0.50,
            'JP':0.50,'LU':0.50,'NL':0.50,'NO':0.50,
            'PO':0.50,'PT':0.50,'SG':0.50,'ES':0.50,
            'SE':0.50,'CH':0.50,'GB':0.50,'US':0.50,
            'BR':0.39,'CZ':0.39,'GR':0.39,'IT':0.39,
            'MX':0.39,'ZA':0.39,'TR':0.39,'UA':0.39}
    LienTwo_map = {'AU':0.18,'AT':0.18,'BE':0.18,
            'CA':0.18,'DK':0.18,'FI':0.18,'FR':0.18,
            'DE':0.18,'HK':0.18,'IE':0.18,'IS':0.18,
            'JP':0.18,'LU':0.18,'NL':0.18,'NO':0.18,
            'PO':0.18,'PT':0.18,'SG':0.18,'ES':0.18,
            'SE':0.18,'CH':0.18,'GB':0.18,'US':0.18,
            'BR':0.13,'CZ':0.13,'GR':0.13,'IT':0.13,
            'MX':0.13,'ZA':0.13,'TR':0.13,'UA':0.13}
    
    bond_map = {'US':0.41}
 
        
    # if it the Recovery rate exists lookup in AAA table
    model_df['S&P Recovery Rate (AAA)'] = model_df['S&P Recovery'].map(new_rr_map)
        #map(dict(new_rr[['S&P Recovery Rating\nand Recovery\nIndicator of\nCollateral Obligations','“AAA”']].values))
    
    # doesn't exist, but first lien, use first lien table
    model_df.loc[pd.isna(model_df['S&P Recovery']) & (model_df['Lien Type']== 'First Lien'),'S&P Recovery Rate (AAA)'] =\
        model_df.loc[pd.isna(model_df['S&P Recovery']) & (model_df['Lien Type']== 'First Lien'),'Issuer Country'].\
        map(LienOne_map)
        #map(dict(lien[['Country Abv','RR']].values))
    
    
    # doesn't exist, but 2nd lien, use 2nd lien table
    model_df.loc[pd.isna(model_df['S&P Recovery']) & (model_df['Lien Type']== 'Second Lien'),'S&P Recovery Rate (AAA)'] = \
        model_df.loc[pd.isna(model_df['S&P Recovery']) & (model_df['Lien Type']== 'Second Lien'),'Issuer Country'].\
        map(LienTwo_map)
        #map(dict(lien[['Country Abv','RR.2nd']].values))
    
    # the bonds
    model_df.loc[pd.isna(model_df['S&P Recovery']) & pd.isna(model_df['Lien Type']),'S&P Recovery Rate (AAA)'] = \
        model_df.loc[pd.isna(model_df['S&P Recovery']) & pd.isna(model_df['Lien Type']),'Issuer Country'].\
        map(bond_map)
        #map(dict(bond_table[['Country Abv.1','RR.1']].values))

    return model_df
#################################################################################
def Specified_Assets(df,col):
    df['S&P CLO Specified Assets'] = df[col].values
    mask = (df['S&P Issuer Rating']=='CC') | \
           (df['S&P Issuer Rating']=='C') | \
           (df['S&P Issuer Rating']=='SD') | \
           (df['S&P Issuer Rating']=='NR') | \
           (df['S&P Issuer Rating']=='D')
    df.loc[mask,'S&P CLO Specified Assets'] = 0   # what about .isna()?
    return df
#################################################################################
def Weighted_Average_Life(clo_df,col):
   
    clo_df['Average Life']  = ((clo_df['Maturity Date'] - pd.Timestamp.today()).dt.days/365).round(2)
    clo_df = Specified_Assets(clo_df,col)
    
    WAL_stat = weighted_average(clo_df,cols=['S&P CLO Specified Assets','Average Life'])
    
    return WAL_stat
#################################################################################
def Weighted_Average_Spread(clo_df,col,libor=.002):
    """
    Per the Virtus Compliance Reports, the Indenture and the Cheat-sheet-Spreadsheet
    the WAS is based on the 'Effective Spread' which is neither the Floating Spread,
    nor the All-in-Rate in the files, but rather derived from the Floating Spread and the 
    Floor vis a vis LIBOR.  It is possible that this needs further refinement, but this
    gets us much closer to the target.
    
    Specified Assets uses all Principal Balance/(Par) less CC/Ca and lower rated deals.
    WAS also uses only Floating Rate Loans, not Fixed Rate Bonds, which are included in the WAC
    
    Args in:
        clo_df:  the dataframe, df
        col:   Which column represents the CLO in question, string
        libor:  latest estimate of LIBOR, float
    Arg out:
        WAS_stat:  The effective WAS, float
    """
   
    clo_df = Specified_Assets(clo_df,col)   # This should be a view, not a copy:  sub CCC excluded
    mask = (clo_df['Asset Type']!='Bond') & (clo_df[col] > 0) & (~clo_df['Spread'].isna())  # Bonds are excluded too
    
    WAS_stat = weighted_average(clo_df.loc[mask],cols=['S&P CLO Specified Assets','Effective Spread'])
    
    return WAS_stat
#################################################################################
def Effective_Spread(df,libor=0.0020):
    """
    The default LIBOR is 20 bps, which was relavant on 6/4/21 
    (to be fair, on 6/4 the 3m LIBOR = 12.7bps, but it was used as
    20bps many places in the examples, so that is the default. 
    Today it is closer to 15bps 6/25/21)
    I'd call the bloomberg api, but then it wouldn't work broadly
    """
    
    def Libor_floor(df,libor = 0.0020):
        df['Libor Floor'] = df['Floor']-libor
        df.loc[df['Libor Floor']<0,'Libor Floor'] = 0
        return df

    df['Effective Spread'] = df['Spread'] + Libor_floor(df,libor = 0.0020)['Libor Floor']

    
    return df
#################################################################################
def Weighted_Average_Coupon(clo_df,col):
    """
    The Test uses WAC + Excess WAFS (currently ~11%)
    The WAC is probably closer to the WAS
    """
    mask = (clo_df['Asset Type']=='Bond')
    WAC_stat = weighted_average(clo_df.loc[mask],[col,'All In Rate'])
    
    return WAC_stat

#################################################################################
def Excess_Weighted_Average_FloatSpread(clo_df,col,minFS):
    """
    This doesn't seem correct.  Is it units or fail?
    """
    mask = (clo_df['Asset Type']!='Bond') & (clo_df[col] > 0) & (~clo_df['Spread'].isna())
    float_Par = clo_df.loc[mask,col].sum()
    
    mask = (clo_df['Asset Type']=='Bond')
    fixed_Par = clo_df.loc[mask,col].sum()
    
    EWAFS = (Weighted_Average_Spread(clo_df,col)-minFS)*float_Par/fixed_Par
    
    return EWAFS
#################################################################################
def Weighted_Average_Life_Test(clo_df,col):
    
    # This is CLO Specific!  It also is the Trigger, not the mapping
    # so this function was barking up the wrong tree.  I.e. you look up
    # the trigger in this map to compare with your test stat
    WAL_value = {'1/9/2021':9.00, '4/9/2021':8.67,'7/9/2021':8.42,'10/9/2021':8.17 ,
        '1/9/2022':7.92,'4/9/2022':7.67,'7/9/2022':7.42 ,'10/9/2022':7.17 ,'1/9/2023':6.92 ,
        '4/9/2023':6.67 ,'7/9/2023':6.42 ,'10/9/2023':6.17 ,'1/9/2024':5.92 ,'4/9/2024':5.67 ,
        '7/9/2024':5.42 ,'10/9/2024':5.17 ,'1/9/2025':4.92 ,'4/9/2025':4.67 ,'7/9/2025':4.42 ,
        '10/9/2025':4.17 ,'1/9/2026':3.92 ,'4/9/2026':3.67 ,'7/9/2026':3.42 ,'10/9/2026':3.17 ,
        '1/9/2027':2.92 ,'4/9/2027':2.67 ,'7/9/2027':2.42 ,'10/9/2027':2.17 ,'1/9/2028':1.92 ,
        '4/9/2028':1.67 ,'7/9/2028':1.42 ,'10/9/2028':1.17 ,'1/9/2029':0.92 ,'4/9/2029':0.67 ,
        '7/9/2029':0.42 ,'10/9/2029':0.17 ,'1/9/2030':0.00 ,'4/9/2030':0.00 ,'7/9/2030':0.00 ,
        '10/9/2030':0.00 ,'1/9/2031':0.00 ,'4/9/2031':0.00 ,'7/9/2031':0.00 ,'10/9/2031':0.00 ,
        '1/9/2032':0.00 ,'4/9/2032':0.00 ,'7/9/2032':0.00 ,'10/9/2032':0.00 ,'1/9/2033':0.00 ,
        '4/9/2033':0.00 ,'7/9/2033':0.00 ,'10/9/2033':0.0}
    
    
    
    WAL_val = pd.DataFrame.from_dict(WAL_value,orient='index',columns=['Value'])
    WAL_val.index = pd.to_datetime(WAL_val.index, dayfirst=False)
    clo_df['Maturity Date'] = pd.to_datetime(clo_df['Maturity Date'], infer_datetime_format=True)
    clo_df = pd.merge_asof(clo_df.sort_values('Maturity Date'), WAL_val, left_on='Maturity Date', 
                          right_on=WAL_val.index, direction='backward', suffixes=['', '_2'])

    clo_df['S&P CLO Specified Assets'] = clo_df[col].values
    mask = (clo_df['S&P Issuer Rating']=='CC') | \
           (clo_df['S&P Issuer Rating']=='C') | \
           (clo_df['S&P Issuer Rating']=='SD') | \
           (clo_df['S&P Issuer Rating']=='NR') | \
           (clo_df['S&P Issuer Rating']=='D')
    clo_df.loc[mask,'S&P CLO Specified Assets'] = 0   # what about .isna()?

    WAL_stat = weighted_average(clo_df,cols=['S&P CLO Specified Assets','Value'])
    
    return WAL_stat
#################################################################################
def percentage_CovLite(model_df,weight_col='Par_no_default'):  
    perC = model_df.loc[model_df['Cov Lite']=='Yes',[weight_col]].sum()/model_df[[weight_col]].sum()
    return perC.values[0]
#################################################################################
def percentage_SecondLien(model_df,weight_col='Par_no_default'):  
    perC = model_df.loc[model_df['Lien Type']=='Second Lien',[weight_col]].sum()/model_df[[weight_col]].sum()
    return perC.values[0]

#################################################################################
def percentage_C(model_df,weight_col='Par_no_default'):  
    perC = model_df.loc[model_df['Adjusted Moodys Rating'].str.match('C'),[weight_col]].sum()/model_df[[weight_col]].sum()
    return perC.values[0]
#################################################################################
def percentage_Caa(model_df,weight_col='CLO 20'):  
    
    perC = model_df.loc[model_df['MDPR'].str.match('C'),[weight_col]].sum()/model_df[[weight_col]].sum()
    return perC.values[0]
#################################################################################
def percentage_CCC(model_df,weight_col='CLO 20'):  
    mask = (model_df['S&P Issuer Rating']=='CC') |  (model_df['S&P Issuer Rating']=='C') |  \
           (model_df['S&P Issuer Rating']=='D') | (model_df['S&P Issuer Rating']=='SD') | \
           (model_df['S&P Issuer Rating']=='NR') #|  (clo_df['S&P Issuer Rating']=='D')
    #perC = model_df.loc[model_df['S&P Issuer Rating'].str.match('C'),[weight_col]].sum()/model_df[[weight_col]].sum()
    perC = model_df.loc[mask,[weight_col]].sum()/model_df[[weight_col]].sum()
    return perC.values[0]

#################################################################################
def Overcollateralization_Stats(df,col,note_dict,format_output=False):
    PrinBal = df[col].sum()
    AB_Denom = note_dict['Class A-1 Notes']+note_dict['Class A-2 Notes']+note_dict['Class B-1 Notes']+\
                note_dict['Class B-2 Notes']
    C_Denom = AB_Denom + note_dict['Class C Notes']
    D_Denom = C_Denom + note_dict['Class D-1 Notes'] + note_dict['Class D-2 Notes']
    E_Denom = D_Denom + note_dict['Class E Notes']
    
    OC_stats_df = pd.DataFrame(np.nan,index=['Class A/B Overcollateralization Ratio',
        'Class C Overcollateralization Ratio',
        'Class D Overcollateralization Ratio',
        'Class E Overcollateralization Ratio',
        'Reinvestment Overcollateralization Ratio'],columns = [col])
    
    OC_stats_df.loc['Class A/B Overcollateralization Ratio',col] = PrinBal/AB_Denom
    OC_stats_df.loc['Class C Overcollateralization Ratio',col] = PrinBal/C_Denom
    OC_stats_df.loc['Class D Overcollateralization Ratio',col] = PrinBal/D_Denom
    OC_stats_df.loc['Class E Overcollateralization Ratio',col] = PrinBal/E_Denom
    OC_stats_df.loc['Reinvestment Overcollateralization Ratio',col] = PrinBal/E_Denom
    
    if format_output:
        OC_stats_df.loc['Class A/B Overcollateralization Ratio'] = \
            (OC_stats_df.loc['Class A/B Overcollateralization Ratio']*100).apply('{:.4f}%'.format)
        OC_stats_df.loc['Class C Overcollateralization Ratio'] = \
            (OC_stats_df.loc['Class C Overcollateralization Ratio']*100).apply('{:.4f}%'.format)
        OC_stats_df.loc['Class D Overcollateralization Ratio'] = \
            (OC_stats_df.loc['Class D Overcollateralization Ratio']*100).apply('{:.4f}%'.format)
        OC_stats_df.loc['Class E Overcollateralization Ratio'] = \
            (OC_stats_df.loc['Class E Overcollateralization Ratio']*100).apply('{:.4f}%'.format)
        OC_stats_df.loc['Reinvestment Overcollateralization Ratio'] = \
            (OC_stats_df.loc['Reinvestment Overcollateralization Ratio']*100).apply('{:.4f}%'.format)
   
    return OC_stats_df
#################################################################################
def prepost_Port_stats(model_df, cols):

    mask = abs(model_df[cols[0]]) > 0
    cstats = Port_stats(model_df.loc[mask], cols[0])
    cstats.rename(columns={'Portfolio Stats':cols[0]},inplace=True)

    if len(cols)>1:
        for c in cols[1:]:
            mask = abs(model_df[c]) > 0
            tstats = Port_stats(model_df.loc[mask], c)
            tstats.rename(columns={'Portfolio Stats':c},inplace=True)
            cstats = cstats.join(tstats)
    
    return cstats

#################################################################################
def all_stats_all_clos(model_df, cols, clo_dict, format_output=False):

    mask = abs(model_df[cols[0]]) > 0
    cstats = Port_stats(model_df.loc[mask], cols[0], format_output)
    ocstats = Overcollateralization_Stats(model_df.loc[mask],cols[0],clo_dict,format_output)
    sstats = SP_CDO_Monitor_Test(model_df.loc[mask],cols[0])
    cstats = cstats.append(ocstats)
    cstats = cstats.append(sstats)
    
    cstats.rename(columns={'Portfolio Stats':cols[0]},inplace=True)

    if len(cols)>1:
        for c in cols[1:]:
            mask = abs(model_df[c]) > 0
            tstats = Port_stats(model_df.loc[mask], c)
            ocstats = Overcollateralization_Stats(model_df.loc[mask],c,clo_dict)
            sstats = SP_CDO_Monitor_Test(model_df.loc[mask],c)
            tstats = tstats.append(ocstats)
            tstats = tstats.append(sstats)
            tstats.rename(columns={'Portfolio Stats':c},inplace=True)
            cstats = cstats.join(tstats)
    
    return cstats
#################################################################################
def build_trigger_tables(trigger_df,clo_list,stats_list):
    trigger_table = pd.DataFrame(np.nan,index=stats_list,columns = clo_list)
    for clo in clo_list:
        for s in stats_list:
            #print('s: ',s,' clo: ',clo)
            try:
                trigger_table.loc[s,clo] = trigger_df.loc[(trigger_df['Title']==s)&
                                                      (trigger_df['CLOName']==clo),'RequirementRaw'].values[0]
            except:
                print(s,' for ',clo, ' DNE')
                #trigger_table.loc[s,clo] = np.nan
    
    return trigger_table
#################################################################################
#def master_test_stats(master_df,weight_col='CLO 2020-20',master_dict):
#################################################################################
def get_triggers(filepathT):
    triggers = pd.read_excel(filepathT)
    ## normalize some of the test names
    triggers.loc[triggers['TestType']=='Coverage Test','Title'] = \
        triggers.loc[triggers['TestType']=='Coverage Test','Title'].map(coverage_map)
    triggers.loc[triggers['TestType']=='Concentration Limitation Test','Title'] = \
        triggers.loc[triggers['TestType']=='Concentration Limitation Test','Title'].str.lstrip("(0123456789abc)/ ")
    triggers.loc[triggers['TestType']=='Concentration Limitation Test','Title'] = \
        triggers.loc[triggers['TestType']=='Concentration Limitation Test','Title'].map(con_stats_map)
    return triggers

#################################################################################
def master_test_stats(df, cols=clo_list, clo_dict= dict_2020_20):

    ## Get All the stats for all CLOs
    all_df = all_stats_all_clos(df, cols, clo_dict, format_output=False)
    all_df.index = all_df.index.map(all_stats_map)
    
    ## Triggers
    triggers = get_triggers(filepathT)
    clo_list = triggers['CLOName'].unique()
    clo_list = sorted(clo_list)
    clo_list = clo_list[11:]+clo_list[0:11]
    
    stats_list = qstats_list + coverage_stats + con_stats # all the triggers
    
    trigger_table = build_trigger_tables(triggers,clo_list,stats_list)
    
    master_test_df = pd.concat([all_df,trigger_table],axis=1,keys=['Result','Trigger']).swaplevel(0,1,axis=1).sort_index(axis=1)
    
    return master_test_df
#################################################################################
## This was older code before the files changed...delete later if not needed
#################################################################################

#path = 'Z:/Shared/Risk Management and Investment Technology/CLO Optimization/'
#file_prefix = 'Portfolio Assets_ Trustee Recon Detail'
#suffixes = [' (CLO 2020-18)',' (CLO 2013-4)',' (CLO 2014-6)',' (CLO 2020-19)',
#            ' (CLO 2020-20)',' (CLO 2014-7)',' (CLO 2015-9)',' (CLO 2020-8r)',
#            ' (CLO 2016-12)',' (CLO 2017-13)',' (CLO 2017-14)',' (CLO 2018-15)',
#            ' (CLO 2019-16)',' (CLO 2019-17)']

def get_trustee_recon(filepath):
    master_df = pd.read_csv(filepath,header=0)
    master_df = master_df.loc[:,~master_df.columns.str.match("Unnamed")]
    master_df.rename(columns={'LoanX Id':'LXID'},inplace=True)
    master_df.set_index('LXID', inplace=True)
    return master_df

