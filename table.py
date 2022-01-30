import pandas as pd
from pandas.core.frame import DataFrame

df = pd.read_excel('table.xlsx')

initialDf = df.iloc[:, 0:2]                       #first model data
initialDf = initialDf.dropna()
initialDf = initialDf.reset_index(drop=True)

num1 = []
for i in range(101, 199, 1):                   # creating range 1.01, 1.02, 1.03 and so on
    num =i/100
    num1.append(num)

newIndexDf = pd.DataFrame()                                  # adding two dataframes together
newIndexDf['newCol'] = num1
mainDf = pd.concat([newIndexDf, initialDf], axis=1).dropna()

#remove duplicates

emptyDf = pd.DataFrame()
removedupDf = pd.DataFrame()
for index, row in df.iterrows():
    emptyDf = row
    row = emptyDf.drop_duplicates()
    #print(row)
    removedupDf = removedupDf.append(row, ignore_index= True)
    emptyDf.iloc[0:0]

columnNew =[]
columnNew = df.columns
removedupDf = removedupDf.reindex(columns = columnNew)
#print(removedupDf)


# get model sequence number
modelDf = pd.DataFrame()
finalDf = pd.DataFrame()
colInd = 0.00
rowInd = 0.00
num = 1.0
num1 =[]
newDf = pd.DataFrame()
rowDf = pd.DataFrame()
allDf = pd.DataFrame()
finalizeDf = []

for row, col in removedupDf.drop(['Work Station'], axis=1).iteritems():
    modelDf = pd.Series(col, name='modelDf')
    rowInd = 0.00
    num1 = []
    for i in modelDf:
        if not pd.isna(i):
            num = (1.0 + colInd) + ((1.0 + rowInd) / 100.0)
            num1.append(num)
            rowInd = rowInd + 1
        else:
            num1.append(0)

    colInd = colInd + 1

    num2 = pd.Series(num1, name='num')
    newDf = pd.merge(modelDf, num2, left_index=True, right_index=True)

    if not newDf.empty:
        finalizeDf.append(newDf)
finalizeDf1 = pd.concat(finalizeDf)



# get workstation
stationDf = df["Work Station"]
finalizeDf1 = pd.merge(stationDf, finalizeDf1,  left_index=True, right_index=True)

# rearrange final dataframe
finalizeDf2= finalizeDf1[finalizeDf1['num'] != 0].sort_values(by= ['num'])
finalizeDf2=finalizeDf2[['num', 'Work Station', 'modelDf']]

#saving data to file

writer = pd.ExcelWriter("result.xlsx", engine='xlsxwriter')
finalizeDf2.to_excel(writer, index=False)
writer.save()

print('Saved data to file')
