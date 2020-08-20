import pybbg
import functions
import datetime
import pandas as pd 
from dateutil.relativedelta import relativedelta

# security = ['58013MEH3 Corp']
security = ['AAPL US 09/18/20 C465 Equity']

#Dataframe for just getting price
roundingDataPrice = pd.DataFrame([
                                 ['PX_LAST','${0:,.2f}',1,'Last Price']
                                 ])

#Dataframe for Equity securities
roundingDataEquities = pd.DataFrame([
                                    ['PX_LAST','${0:,.2f}',1,'Last Price'],
                                    ['CUR_MKT_CAP','${:,.0f}',1000,'Curr Mkt Cap'],
                                    ['PX_VOLUME','{0:,.0f}',1,'Volume'],
                                    ])

#Dataframe for CORP securities, bonds, UST, etc.
roundingDataBonds = pd.DataFrame([
                                 ['PX_LAST','${0:,.2f}',1,'Last Price'],
                                 ['BLP_SPRD_TO_BENCH_BID','{:,.1f}',1,'BBG Bid Sprd To Benchmark'], 
                                 ['BLP_SPRD_TO_BENCH_ASK','{:,.1f}',1,'BBG Ask Sprd To Benchmark'],   
                                 ])

#main BDH
def bdh(securities, roundingData, histDate, today):
    try:
        pipeline = pybbg.Pybbg()
        data = pd.DataFrame(pipeline.bdh(securities, roundingData, histDate, today))
        return(data)
    except Exception as e: print(e) 


today = datetime.datetime.today()
histDate = today + datetime.timedelta(days=-365)

outputTable = bdh(security, roundingDataPrice[0], histDate, today)
functions.printFrame(security, outputTable, roundingDataPrice)
functions.plotGraph(outputTable)
