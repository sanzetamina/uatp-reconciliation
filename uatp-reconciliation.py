import pandas as pd
import glob

# to be able to see the full cols/row of dataframes on the terminal
pd.set_option("display.max_rows", None)

# read all the Excel files in the directory as input, exclude any files with 'output' in their names - because those are the ones we'll be producing
input_files = list(set(glob.glob("*.xls*")) - set(glob.glob("*utput*")))

print("---- INPUT FILES ----")
print(input_files)
print("---------------------")

appended_data = []
for f in input_files:
    data = pd.read_excel(f)
    # store DataFrame in list
    appended_data.append(data)
appended_data = pd.concat(appended_data)

# initialise a new dataframe loaded with data from files in directory
df = pd.DataFrame(appended_data)

# To prevent problems further down the line, filling-in empty cells with a 0 value
df.fillna("NIL", inplace=True)

# Formatting fields
df["TRANSACTION NUMBER"] = df["TRANSACTION NUMBER"].apply("{:0>13}".format)
df["TRANSACTION NUMBER"] = df["TRANSACTION NUMBER"].str.replace(
    r"(?<=^.{3})", r"-", regex=True
)

# print loaded dataframe
print("---- Imported Excel Data ----")
print(df)

# dataframe col types
print(" --- COL DTYPES ----")
print(df.dtypes)

# dataframe col types
print(" --- MODIFIED DTYPES ----")
print(df.dtypes)

######################################### CHOOSE ONE THE THREE ALTERNATIVES HERE #########################################
### ALTERNATIVE 1 ###
# create pivot table by TICKET then PNR rows, and sort by column Total ascending
# pivot_table = pd.pivot_table(df, index=['TRANSACTION NUMBER', 'CUSTOMER REFERENCE'], columns='TRANSACTION TYPE',
#                              values='BILLING VALUE', fill_value='.', aggfunc=sum, margins=True, margins_name='Total', sort=True)

### ALTERNATIVE 2 ###
# create pivot table by PNR then TICKET rows, and sort by column Total ascending
pivot_table = pd.pivot_table(
    df,
    index=["CUSTOMER REFERENCE", "TRANSACTION NUMBER"],
    columns="TRANSACTION TYPE",
    values="BILLING VALUE",
    fill_value=".",
    aggfunc=sum,
    margins=True,
    margins_name="Total",
    sort=True,
)

### ALTERNATIVE 3 ###
# create pivot table by TICKET ONLY rows, and sort by column Total ascending
# pivot_table = pd.pivot_table(df, index = ['TRANSACTION NUMBER'], columns = 'TRANSACTION TYPE', values='BILLING VALUE', fill_value='.', aggfunc=sum, margins=True, margins_name='Total', sort=True)
# .sort_values(by=['Total'], ascending=True)
##########################################################################################################################

# remove row Total from Pivot table, we keep though the Total column
pivot_table.drop("Total", axis=0, inplace=True)

print("---- PIVOT ----")
print(pivot_table)

# Convert pivot_table to dataframe
dfPivot = pivot_table.reset_index()

# Sort dfPivot by Total ascending
dfPivot.sort_values(by=["Total"], ascending=True, inplace=True)

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

# Implement a 2nd round of re-processing, let's call it "pnr_grouped", this time taking as source the "dfOutstandingTrxs" dataframe
# since there are many instances of PNRs with more than one ticket, they end up not being considered as settled (many false positives)
# by doing this additional round, we hope we can reduce the output of outstanding trx by removingn the false positives which are actually settled (via exchanged tickets)
# create new pivot table having dfOutstandingTrxs as input
pnr_grouped_pivot_table = pd.pivot_table(
    dfOutstandingTrxs,
    index=["CUSTOMER REFERENCE"],
    # columns=[],
    values="Total",
    fill_value=".",
    aggfunc=sum,
    margins=True,
    margins_name="Total",
)

# Format the "Total" column as a float with two decimal places
# pnr_grouped_pivot_table["Total"] = pnr_grouped_pivot_table["Total"].map("{:.2f}".format)

# print pnr_grouped_pivot_table:
print("----- pnr_grouped PIVOT ------")
print(pnr_grouped_pivot_table.round(2))

# Convert pnr_grouped_pivot_table to dataframe
dfPnrGrouped = pnr_grouped_pivot_table.reset_index()

# Format the "Total" column as a float with two decimal places
dfPnrGrouped["Total"] = dfPnrGrouped["Total"].round(2)

# Sort dfPnrGrouped by Total column in ascending order
dfPnrGrouped.sort_values(by="Total", ascending=True, inplace=True)

# split the dataframe "dfPnrGrouped" onto two separate ones, for Settled PNRs and Outstanding PNRs to later place them into different Excel sheets
dfSettledPNRs = dfPnrGrouped[dfPnrGrouped.Total == 0]
dfOutstandingPNRs = dfPnrGrouped[dfPnrGrouped.Total != 0]

# Sort dfOutstandingPNRs by Total column in ascending order
dfOutstandingPNRs.sort_values(by="Total", ascending=True, inplace=True)

# Print dfSettledPNRs:
print("----- Settled PNRs - dfSettledPNRs ------")
print(dfSettledPNRs)

# Print dfOutstandingPNRs:
print("----- Outstanding PNRs - dfOutstandingPNRs ------")
print(dfOutstandingPNRs)

# Write output excel
with pd.ExcelWriter("Output.xlsx") as writer:
    df.to_excel(writer, "UATP Source")
    dfPivot.to_excel(writer, "Pivot")
    dfSettledTrxs.to_excel(writer, "Settled Trxs")
    dfOutstandingTrxs.to_excel(writer, "Outstanding Trxs")
    dfSettledPNRs.to_excel(writer, "Settled PNRs")
    dfOutstandingPNRs.to_excel(
        writer, "Outstanding PNRs", index=False, float_format="%.2f"
    )


# TODO
# row index on 'UATP Source' doesnt match the row index on the subsequent tabs
# ticket num. duplicates on subsequent tabs (to do with whether to group by PNR or not) - maybe do vlookup instead
