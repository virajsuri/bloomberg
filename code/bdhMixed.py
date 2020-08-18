import pybbg
import functions
import datetime
import pandas as pd 
from dateutil.relativedelta import relativedelta

equity = ['MSFT US Equity']

# [FLDS,FORMATTING,MULTIPLIER,LONG NAME]
roundingDataEquities = pd.DataFrame([
                                    ['PX_LAST','${0:,.2f}',1,'Last Price'],
                                    ['CUR_MKT_CAP','${:,.0f}',1000,'Curr Mkt Cap'],
                                    ['PX_VOLUME','{0:,.0f}',1,'Volume'],
                                    ])

def bdh_mixed(securities, roundingData, histDate, today, periodicity):
        pipeline = pybbg.Pybbg()
        data = pipeline.bdh(securities,roundingData[0], histDate, today, periodicity,
            overrides=dict(
                CALENDAR_CONVENTION=1 # calendar convention 'calendar' flds rk408
            ),
            other_request_parameters=dict(
                periodicityAdjustment='CALENDAR',
                nonTradingDayFillMethod='PREVIOUS_VALUE',
                returnRelativeDate=True
            ),
            move_dates_to_period_end=True
        ).iloc[::-1]
        return(data)


def mainMethod():
    today = datetime.datetime.today()
    histDate = today + datetime.timedelta(days=-365)

    outputTable = bdh_mixed(equity, roundingDataEquities, histDate, today, 'MONTHLY')
    # print(outputTable)
    functions.printFrame(equity, outputTable, roundingDataEquities)

mainMethod()