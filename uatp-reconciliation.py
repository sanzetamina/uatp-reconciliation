import pandas as pd
import glob

# to be able to see the full cols/row of dataframes on the terminal
pd.set_option('display.max_rows', None)

# read all the Excel files in the directory as input, exclude any files with 'output' in their names - because those are the ones we'll be producing
input_files = list(set(glob.glob("*.xls*")) - set(glob.glob("*utput*")))

print('---- INPUT FILES ----')
print(input_files)
print('---------------------')

appended_data = []
for f in input_files:
    data = pd.read_excel(f)
    # store DataFrame in list
    appended_data.append(data)
appended_data = pd.concat(appended_data)

# initialise a new dataframe loaded with data from files in directory
df = pd.DataFrame(appended_data)

# To prevent problems further down the line, filling-in empty cells with a 0 value
df.fillna('NIL', inplace=True)

# Formatting fields
df['TRANSACTION NUMBER'] = df['TRANSACTION NUMBER'].apply('{:0>13}'.format)
df['TRANSACTION NUMBER'] = df['TRANSACTION NUMBER'].str.replace(
    r'(?<=^.{3})', r'-', regex=True)

# print loaded dataframe
print("---- Imported Excel Data ----")
print(df)

# dataframe col types
print(" --- COL DTYPES ----")
print(df.dtypes)

# dataframe col types
print(" --- MODIFIED DTYPES ----")
print(df.dtypes)

# create pivot table by TICKET then PNR rows, and sort by column Total ascending
pivot_table = pd.pivot_table(df, index=['TRANSACTION NUMBER', 'CUSTOMER REFERENCE'], columns='TRANSACTION TYPE',
                             values='BILLING VALUE', fill_value='.', aggfunc=sum, margins=True, margins_name='Total', sort=True)

# create pivot table by PNR then TICKET rows, and sort by column Total ascending
# pivot_table = pd.pivot_table(df, index = ['CUSTOMER REFERENCE', 'TRANSACTION NUMBER'], columns = 'TRANSACTION TYPE', values='BILLING VALUE', fill_value='.', aggfunc=sum, margins=True, margins_name='Total', sort=True)

# create pivot table by TICKET ONLY rows, and sort by column Total ascending
# pivot_table = pd.pivot_table(df, index = ['TRANSACTION NUMBER'], columns = 'TRANSACTION TYPE', values='BILLING VALUE', fill_value='.', aggfunc=sum, margins=True, margins_name='Total', sort=True)
# .sort_values(by=['Total'], ascending=True)

# remove row Total from Pivot table, we keep though the Total column
pivot_table.drop('Total', axis=0, inplace=True)

print("---- PIVOT ----")
print(pivot_table)

# Convert pivot_table to dataframe
dfPivot = pivot_table.reset_index()

# Sort dfPivot by Total ascending
dfPivot.sort_values(by=['Total'], ascending=True, inplace=True)

# VLOOKUP PNR
# dfPivot['PNR']=df['CUSTOMER REFERENCE'].apply(lambda x: df['CUSTOMER REFERENCE'])
# df.merge(dfPivot, on='CUSTOMER REFERENCE', how='left')

print("----- DATAFRAME PIVOT ------")
print(dfPivot)

# split the dataframe onto two separate ones, for Settled transacitons and Outstanding ones to later place them into different Excel sheets
dfSettledTrxs = dfPivot[dfPivot.Total == 0]
dfOutstandingTrxs = dfPivot[dfPivot.Total != 0]

print("--- SETTLED TRXS ----")
print(dfSettledTrxs)

print("--- OUTSTANDING TRXS ----")
print(dfOutstandingTrxs)

# Write output excel
with pd.ExcelWriter('Output.xlsx') as writer:
    df.to_excel(writer, 'UATP Source')
    dfPivot.to_excel(writer, 'Pivot')
    dfSettledTrxs.to_excel(writer, 'Settled Trxs')
    dfOutstandingTrxs.to_excel(writer, 'Outstanding Trxs')

# TODO
# row index on 'UATP Source' doesnt match the row index on the subsequent tabs
# ticket num. duplicates on subsequent tabs (to do with whether to group by PNR or not) - maybe do vlookup instead
