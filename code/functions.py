import matplotlib.pyplot as plt


#show plot graph with one security and one FLD
def plotGraph(outputTable):
    print("Plotting graph")
    plt.figure(figsize=(10,4))
    plt.plot(outputTable, color='black',linestyle='solid')
    plt.title(outputTable.columns[0])
    plt.tight_layout()
    plt.savefig(outputTable.columns[0]+".jpg", dpi=300)
    plt.show()


#print function for BDH reqs in the console
def printFrame(equity, table, roundingData):
    print(equity[0])
    for counterIndex, numOfFieldsIndex, in enumerate(table):

        fldName = roundingData[0][counterIndex]
        formatStyle = roundingData[1][counterIndex]
        multiplier = roundingData[2][counterIndex]
        longName = roundingData[3][counterIndex]
       
        for x in range(len(table[numOfFieldsIndex])):
            print(table[numOfFieldsIndex].index[x].strftime("%Y-%m-%d "),end="")
            print(longName+": ",end="")
            value = table[numOfFieldsIndex][x]
            print(formatStyle.format(multiplier*value),end="")
            table[numOfFieldsIndex][x] = (multiplier*value)
            print()
 
        print()
    # return(table) 