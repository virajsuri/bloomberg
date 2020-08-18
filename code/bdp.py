import pybbg
import pandas as pd 

# [FLDS,FORMATTING,MULTIPLIER,LONG NAME]

ustBonds = ['GT2 Govt','GT3 Govt','GT5 Govt','GT7 Govt','GT10 Govt','GT30 Govt']
roundingDataUST = pd.DataFrame([['yld_ytm_mid','{0:.2%}',.01,'YTW'],
                            ['yld_chg_cls_ask','{0:.1f} bps',100,'YLD Chg 1D'],
                            ['yld_chg_mtd_mid_net','{0:.1f} bps',100,'Chg MTD'],
                            ['yld_chg_ytd_mid_net','{0:.1f} bps',1,'Chg YTD']])

equityIndicies = ['INDU Index','SPX Index','CCMP Index']
roundingDataEquities = pd.DataFrame([
                                    ['px_last','${0:.2f}',1,'Last Price'],
                                    ['chg_pct_1d','{0:.2%}',.01,'Price Chg 1D'],
                                    ['chg_pct_mtd','{0:.2%}',.01,'Price Chg MTD'],
                                    ['chg_pct_ytd','{0:.2%}',.01,'Price Chg YTD'],
                                    ])

# Returns a pandas DF given securities and fields
def test_bdp(securities, roundingData):
        try:
            pipeline = pybbg.Pybbg()
            data = pd.DataFrame(pipeline.bdp(securities, roundingData[0]))
        except Exception as e: print(e)
        return(data) 

#Returns the pandas DF with the multipliers applied
def printFrame(table, roundingData):
    for index in table:
        print(index)
        for x in range(len(roundingData)):
            value = (table[index][x])
            
            formatStyle = roundingData[1][x]
            multiplier = (roundingData[2][x])
            fieldName = roundingData[3][x]
            
            print(fieldName+" "+formatStyle.format(value*multiplier))
            table[index][x] = (value*multiplier)
        print()
    return(table)    

#main method to get data
def getData(securities, roundingData):
    outputTable = test_bdp(securities, roundingData)
    formattedOutputTable = printFrame(outputTable, roundingData)
    print(formattedOutputTable)

# getData(ustBonds, roundingDataUST)
getData(equityIndicies, roundingDataEquities)
        