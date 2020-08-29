import pybbg
from datetime import date
import datetime
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

tickers = [
    'NCOVUSCA Index',
    'NCOVUSDE Index',
    'NCOVUSRE Index',
    'NCOVCASE Index',
    'NCOVDEAT Index',
    'NCOVRECO Index',
    'SCOVCACA Index',
    'SCOVFLCA Index',
    'SCOVNYCA Index',
    'SCOVTXCA Index',
    'SCOVNJCA Index',
    'MTAFINTL Index',
    'MTAFINMN Index',
    'OPENCOUS Index',
    'OPENCIDV Index',
    'OPENCIWA Index',
    'OPENCINY Index',
    # 'MOVTUSWA Index',
    # 'MOVTUSCH Index',
    # 'MOVTUSBO Index',
    # 'MOVTUSMI Index',
    # 'MOVTUSPH Index',
    # 'MOVTUSHO Index',
    # 'MOVTUSLO Index',
    # 'MOVTUSPI Index',
    # 'MOVTUSSE Index',
    # 'MOVTUSSF Index',
    # 'TSATTPCY Index',
    # 'TSATTPPY Index',
    # 'GSDEDUSR Index',
    # 'GSDEDUSP Index',
    # 'USHMEWTO Index',
    # 'USHMHWTO Index',
    # 'USHMLBTO Index',
    # 'RSPR Index',
    # 'INJCJC Index',
    # 'INJCSP Index',
]

roundingData = pd.DataFrame([['PX_LAST']])

compName = pd.DataFrame([['LONG_COMP_NAME']])

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
        )
        return(data)

def bdp(securities, roundingData):
        try:
            pipeline = pybbg.Pybbg()
            data = pd.DataFrame(pipeline.bdp(securities, roundingData[0]))
        except Exception as e: print(e)
        return(data) 

def mainMethod():
    #get dates
    endDate = datetime.date.today()
    startDate = date(2020,3,1)    
    
    #get covid data
    output_table = bdh_mixed(tickers, roundingData, startDate, endDate, 'WEEKLY')
    # print(output_table)

    #get long comp names
    comp_names = bdp(tickers, compName)
    # print(comp_names)

    #save to excel
    print("Saving to Excel")
    writer = pd.ExcelWriter('../data/covidData.xlsx', engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
    comp_names.to_excel(writer, sheet_name='COMP_NAMES')
    output_table.to_excel(writer, sheet_name='DATA')
    writer.save()
    writer.close()

    #plot and save as PDF
    pdf = PdfPages("../data/covidoutput.pdf")
    for columns in comp_names.columns:
        title = comp_names[columns][0]
        print("Plotting "+title)

        fig = plt.figure(figsize=(9,5))
        plt.plot(output_table[columns])
        plt.title(title)
        # plt.ylim(bottom=0)

        axes = plt.gca()
        axes.yaxis.grid()
        plt.tight_layout()
        plt.savefig("../data/covid/"+title+".jpg", dpi=300, bbox_inches='tight')
        pdf.savefig(fig)
        plt.close()
        
    pdf.close()
    

mainMethod()