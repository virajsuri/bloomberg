import pybbg
from datetime import date
import datetime
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.ticker import FuncFormatter

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
    'MOVTUSWA Index',
    'MOVTUSCH Index',
    'MOVTUSBO Index',
    'MOVTUSMI Index',
    'MOVTUSPH Index',
    'MOVTUSHO Index',
    'MOVTUSLO Index',
    'MOVTUSPI Index',
    'MOVTUSSE Index',
    'MOVTUSSF Index',
    'TSATTPCY Index',
    'TSATTPPY Index',
    'USHMEWTO Index',
    'USHMHWTO Index',
    'USHMLBTO Index',
    'RSPR Index',
    'INJCJC Index',
    'INJCSP Index'
]

def bdh_mixed(securities, field, histDate, today, periodicity):
        pipeline = pybbg.Pybbg()
        data = pipeline.bdh(securities, field, histDate, today, periodicity,
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
            data = pd.DataFrame(pipeline.bdp(securities, roundingData))
        except Exception as e: print(e)
        return(data) 

def saveToExcel(comp_names, output_table):
    print("Saving to Excel")
    writer = pd.ExcelWriter('../data/covidData.xlsx', engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
    output_table.to_excel(writer, sheet_name='DATA')
    comp_names.to_excel(writer, sheet_name='COMP_NAMES')
    writer.save()
    writer.close()

def plot(comp_names, output_table):
    pdf = PdfPages("../data/covidoutput.pdf")
    for columns in output_table.columns:
        title = comp_names[columns][0]

        #title processing
        data_pt_found = False
        data_pt_counter = 1
        while data_pt_found == False:
            try:
                recentDataPt = f'{int(output_table[columns][len(output_table)-data_pt_counter]):,}'
                data_pt_found = True
            except Exception as e: 
                # print(e)
                data_pt_counter+=1

        mostrecent = " (Most Recent: " +recentDataPt+")"
        print("Plotting "+title+mostrecent)

        #plot
        fig = plt.figure(figsize=(9,5))
        plt.plot(output_table[columns])
        plt.title(title+mostrecent)

        #add gridlines
        axes = plt.gca()
        axes.yaxis.grid()

        #format y axis
        axes.get_yaxis().set_major_formatter(FuncFormatter(lambda x, p: format(int(x), ',')))

        #process and save
        plt.tight_layout()
        plt.savefig("../data/covid/"+title+".jpg", dpi=300, bbox_inches='tight')
        pdf.savefig(fig)
        plt.close()
        
    pdf.close()


def main():
    #get dates
    endDate = datetime.date.today()
    startDate = date(2020,3,1)    
    
    #get covid data
    output_table = bdh_mixed(tickers, 'PX_LAST', startDate, endDate, 'WEEKLY')
    # print(output_table)

    #get long comp names
    comp_names = bdp(tickers, 'LONG_COMP_NAME')
    # print(comp_names)

    # saveToExcel(comp_names, output_table)
    plot(comp_names, output_table) 
    
main()