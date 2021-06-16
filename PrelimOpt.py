'''
These were functions first derived in the CLO Optimization problem.

The raise_WAS and lower_WARF functions are more in the DP style, but
somewhat limited in use as a result.  CLOOpt_small was an attempt
to limit the total trade optimization problem to a certain number
of trades and precedes the TradeMinimization Formulation.
'''



import numpy as np  
import pandas as pd
#import datetime as dt
#import blpapi
#from xbbg import blp
from pyxll import xl_func
from pulp import *

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

############################################################################################
@xl_func("dataframe<index=True>, str, dict<str, float>, dict<str, int>, dataframe<index=True>, str: object")
def CLOOpt_Small(model_dfO,userVar,keyConstraints,otherConstraints,seedTrades,probName):  #,targets, ,exclusionCrit, ,exclusionTrades
    """
    This uses the PuLP optimizer whose objective is to try to maximize 
    desirability given a set of constraints. You control which stats you
    want to maximize (e.g. WAS, (-)WARF, etc) by choosing higher weights
    in the desirability for the specified stat(s)
    
    Args in:
        model_df: dataframe or object creating with the universe ands stats
        keyStats: the stats which will be used in desirability and constraints
        weights: weights in desirability
        highLow: specifies whether the higher stat is good (e.g WAS) or low is good (e.g. WARF)
        contraints: Constraints for the keyStats
        targets: (future use)
        Cash_to_Spend: if there is cash to spend or raise (-) then specify
        probName: to name the problem to refer to the .lp later
    Args out:
        model_df: NewPort, CashDelta, Trade columns added with the solution if feasible
    
    """
    
    # from constraintDict we should build the constraints
    # the intention is that this dict will be built from
    # user determined constraints and weighting of the
    # objective
    model_df = model_dfO.copy()
    currStats = Port_stats(model_df,weight_col='Par_no_default',format_output=False)
    currStats = dict(zip(currStats.index,currStats.values))
    print(currStats)
    
    # this would be better as a dictionary in case the sheet changes order
    Cash_to_Spend = otherConstraints['Cash to spend/raise']
    PBLim = otherConstraints['Par Build(+) Loss(-) Limit']
    upperTradable = otherConstraints['Max trade size (on buys)']
    maxTrades = otherConstraints['Max # of new loans']   # not using atm
    print('Cash_to_Spend: ',Cash_to_Spend,' PBLim: ',PBLim,' upperTradable: ',upperTradable)
      
        
    # Let's cull the list down to a smaller list to optimize on
    # Select the Max#Trades with the highest Desirability score in addition
    # to the Current Portfolio and let the optimizer use those only
    
    #model_df = model_df[model_df]
    # top trades not currently in portfolio
    #model_df.loc[model_df['Categorical']=='Outside'].sort_values(by='Desirability', ascending=False)[0:maxTrades-1]
    newLoans = model_df.loc[model_df['Categorical']=='Outside'].nlargest(maxTrades,'Desirability').index
    model_df = model_df.loc[(model_df['Categorical']=='Current') | 
               model_df.index.isin(newLoans)]
        
    WARFTest = keyConstraints['Adj. WARF NEW']
    WARFcp = currStats['Max Moodys Rating Factor Test (NEW WARF)'][0]
    WARFdelta = WARFTest - WARFcp
    RecoveryTest = keyConstraints['S&P Recovery Rate (AAA)']
    RRcp = currStats['Min S&P Recovery Rate Class A-1a'][0]
    RRdelta = RecoveryTest - RRcp
    DiversityTest = -100
    print('WARFTest: ',WARFTest,' WARFcp: ',WARFcp,' RecoveryTest: ',RecoveryTest,' RRcp: ',RRcp)

    TotalPar = currStats['Total Portfolio Par (excl. Defaults)'][0]
    #ParDenom = currStats[13]+PBLim # this is the lower limit of the new Par amt, use for % constraints
    SubC_Constr = keyConstraints['C_or_Less']*(TotalPar+PBLim) - currStats['Percent CCC & lower'][0]*TotalPar/100
    Lien_Constr = keyConstraints['Lien']*(TotalPar+PBLim) - currStats['Percent 2nd Lien'][0]*TotalPar/100
    Sub80_Constr = keyConstraints['Sub80']*(TotalPar+PBLim) - currStats['Percent Sub80'][0]*TotalPar/100
    Sub90_Constr = keyConstraints['Sub90']*(TotalPar+PBLim) - currStats['Percent Sub90'][0]*TotalPar/100
    Cov_Constr = keyConstraints['CovLite']*(TotalPar+PBLim) - currStats['Percent CovLite'][0]*TotalPar/100
    
    print('Lien: ',Lien_Constr,' Cov: ',Cov_Constr,' SubC: ',SubC_Constr,' Sub80: ',Sub80_Constr)
    
    # Create the 'prob' variable to contain the problem data
    prob = LpProblem(probName,LpMaximize)   # <- Max attractiveness, could be just Spread

    # Creates a list of the Features to use in problem
    Trades = model_df.index
    Trades = Trades.unique()  # edit 6/4/21 fixed this in orig df creation, left here JIC

    # A dictionary of the attractiveness of each of the Loans is created
    userMax = dict(zip(model_df[userVar].index,model_df[userVar].values))
    desirability = dict(zip(model_df['Desirability'].index,model_df['Desirability'].values))

    # A dictionary of the Rating Factor in each of the Loans is created
    WARF = dict(zip(model_df['Adj. WARF NEW'].index,model_df['Adj. WARF NEW'].values))

    # A dictionary of the Recovery Rate in each of the Loans is created
    WARR = dict(zip(model_df['S&P Recovery Rate (AAA)'].index,model_df['S&P Recovery Rate (AAA)'].values)) #,'MC Div Score'
    
    # dictionaries for Lien Type, Low Ratings, Cov Lite & Sub 80/90 Priced loans
    LienTwo = dict(zip(model_df['Lien'].index,model_df['Lien'].values))
    CovLite = dict(zip(model_df['CovLite'].index,model_df['CovLite'].values))
    SubC    = dict(zip(model_df['C_or_Less'].index,model_df['C_or_Less'].values))
    SubEighty = dict(zip(model_df['Sub80'].index,model_df['Sub80'].values))
    SubNinety = dict(zip(model_df['Sub90'].index,model_df['Sub90'].values))
    

    # A dictionary of the current port positions in each of the Loans is created
    CP = dict(zip(model_df['Current Portfolio'].index,model_df['Current Portfolio'].values)) #,'CP'
    Ask = dict(zip(model_df['Ask'].index,model_df['Ask'].values))
    Bid = dict(zip(model_df['Bid'].index,model_df['Bid'].values))
    #Mid = dict(zip(model_df['Mid'].index,model_df['Mid'].values))
    #APP = dict(model_port['Blended Actual Purchase Prices'])
    #Ind = dict(zip(model_df['Ind'].index,model_df['Ind'].values))  # this was just a col of 1s, like I but vector, not needed

    # A dictionary of the fibre percent in each of the Loans is created
    mcDiversity = dict(zip(model_df['MC Div Score'].index,model_df['MC Div Score'].values)) 
    
    # create the loan-wise upper and lower limits
    UB = CP.copy()
    LB = CP.copy()
    for k  in CP:
        LB[k] = max(-CP[k],-upperTradable)  # no shorting and limited sell amt
        UB[k] = upperTradable
    
    # seed trades should be set like x.lowBound = seedAmt, where x is the LXID variable (could set lb=ub=seedAmt)
    # likewise loans to not buy x.upBound = 0, and to not sell x.lowBound = CP_i (or simply drop from DF)
    if ~seedTrades.isnull().values.all():
        for i in seedTrades.index:
            LB[i] = seedTrades.loc[i].values[0]
            UB[i] = seedTrades.loc[i].values[0]

    # Can't buy unAttractive loans
    for i in model_df.loc[(model_df['Attractiveness']==1) | (model_df['Attractiveness']==2) ].index:
        UB[i] = 0
        
    # Can't sell very Attractive loans *we own*
    # for i in model_df.loc[(model_df['Attractiveness']==5) & (model_df['Current Portfolio']!=0) ].index:
    
    # Can't sell very Attractive loans
    for i in model_df.loc[(model_df['Attractiveness']==5) ].index:
        LB[i] = 0

    # This limit the sells to the amount in current portfolio and up to amt tradeable
    # I need the -cp_i <= t_i <= trade_size_limit (e.g. 1e6)
    trades = [LpVariable(format(i), lowBound = LB[i],  upBound = UB[i]) for i in Trades]
    
    # The objective function is added to 'prob' first
    #prob += lpSum([desirability[i] * t for t, i in zip(trades,Trades)]), "Total Desirability of Loan Portfolio"
    prob += lpSum([userMax[i] * t for t, i in zip(trades,Trades)]), "Max/Min of " + userVar + " of Loan Portfolio"
    
    # First the practical constraints are added to 'prob' (self-funding, parburn, etc)
    prob += lpSum([ Bid[i]/100 * t for t, i in zip(trades,Trades)]) <= Cash_to_Spend , "Self-funding Bid"
    prob += lpSum([ Ask[i]/100 * t for t, i in zip(trades,Trades)]) <= Cash_to_Spend , "Self-funding Ask"
    #prob += lpSum([ Mid[i]/100 * t for t, i in zip(trades,Trades)]) <= Cash_to_Spend , "Self-funding Mid"
    # I think this needs to be Bid for CP and Ask for loan
    prob += lpSum([((100-Bid[i])/100 * t) for t, i in zip(trades,Trades)]) >= PBLim, "Par Burn Limit"
    #prob += lpSum([((100-Mid[i])/100 * t) for t, i in zip(trades,Trades)]) >= PBLim, "Par Burn Limit"
    # still kind of weird that this is needed, must be in corner solution
    prob += lpSum([((1+(100-Bid[i])/100) * t) for t, i in zip(trades,Trades)]) >= 0, "Must use cash raised"
    #prob += lpSum([((1+(100-Mid[i])/100) * t) for t, i in zip(trades,Trades)]) >= 0, "Must use cash raised"

    # this doesn't work! 
    #prob += lpSum([Ind[i] for i in Trades]) <= maxTrades, "Maximum Number of Trades"
    #prob += lpSum([ t for t, i in zip(trades,Trades)]) <= maxTrades*upperTradable, "Maximum Number of Trades"   
    #prob += lpSum([ t - CP[i] for t, i in zip(trades,Trades)]) <= 2*maxTrades*upperTradable, "Maximum Number of Trades"   
    
    # then the Test Condition Hard constriants, WARF,RR, Div, etc
    prob += lpSum([WARF[i] * t for t, i in zip(trades,Trades)]) <= WARFdelta, "WARF Test"
    prob += lpSum([WARR[i] * t for t, i in zip(trades,Trades)]) >= RRdelta, "Recovery Test"
    
    # need to derive a better representation of this constraint
    prob += lpSum([mcDiversity[i] * t for t, i in zip(trades,Trades)]) >= DiversityTest, "Diversity Test (simplified)"
    
    #
    prob += lpSum([CovLite[i] * t for t, i in zip(trades,Trades)]) <= Cov_Constr, "Cov Test"
    prob += lpSum([SubC[i] * t for t, i in zip(trades,Trades)]) <= SubC_Constr, "Sub C Test"
    prob += lpSum([SubEighty[i] * t for t, i in zip(trades,Trades)]) <= Sub80_Constr, "Sub 80 Test"
    prob += lpSum([SubNinety[i] * t for t, i in zip(trades,Trades)]) <= Sub90_Constr, "Sub 90 Test"
    prob += lpSum([LienTwo[i] * t for t, i in zip(trades,Trades)]) <= Lien_Constr, "2nd Lien Test"

    # The problem data is written to an .lp file
    prob.writeLP(probName+".lp")

    # The problem is solved using PuLP's choice of Solver
    prob.solve()

    # The status of the solution is printed to the screen
    print("Status:", LpStatus[prob.status])

    # Each of the variables is printed with it's resolved optimum value
    # so this would be new portfolio and new to derive trades by comparing to old
    for v in prob.variables():
        #print(v.name, "=", v.varValue)
        model_df.loc[v.name,'NewPort'] = model_df.loc[v.name,'Par_no_default'] + v.varValue * \
            (1+(100-model_df.loc[v.name,'Ask'])/100 if v.varValue > 0 else 
            1+(100-model_df.loc[v.name,'Bid'])/100 )
        model_df.loc[v.name,'CashDelta'] = v.varValue
        model_df.loc[v.name,'Trade'] = 'Buy' if v.varValue > 0 else 'Sale' if v.varValue < 0 else np.nan
    
    return model_df