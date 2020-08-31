import pybbg
from datetime import date
import datetime
import pandas as pd
from pprint import pprint

tickers = [
    # 'NCOVUSCA Index',
    # 'NCOVUSDE Index',
    # 'NCOVUSRE Index',
    # 'NCOVCASE Index',
    # 'NCOVDEAT Index',
    # 'NCOVRECO Index',
    # 'SCOVCACA Index',
    # 'SCOVFLCA Index',
    # 'SCOVNYCA Index',
    # 'SCOVTXCA Index',
    # 'SCOVNJCA Index',
    # 'MTAFINTL Index',
    # 'MTAFINMN Index',
    # 'OPENCOUS Index',
    # 'OPENCIDV Index',
    # 'OPENCIWA Index',
    # 'OPENCINY Index',
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
    # 'USHMEWTO Index',
    # 'USHMHWTO Index',
    # 'USHMLBTO Index',
    'RSPR Index',
    'INJCJC Index',
    'INJCSP Index'
]

field = 'PX_LAST'

compName = 'LONG_COMP_NAME'

def bdp(securities, roundingData):
        try:
            data = pd.DataFrame(pipeline.bdp(securities, roundingData))
        except Exception as e: print(e)
        return(data) 


#open pipeline
pipeline = pybbg.Pybbg()

comp_names = bdp(tickers, 'LONG_COMP_NAME')
# print(comp_names)

for columns in comp_names:
    print(comp_names[columns])