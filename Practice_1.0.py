import pandas as pd
import openpyxl
import seaborn as sns

pd.set_option('display.width', 450)
pd.set_option('display.max_columns', 20)

wb = openpyxl.load_workbook(filename='GL 1000.xlsx')

# Obtain all the sheet names in the workbook.
wb.sheetnames

# Since the file is .xlsx need to use the openpyxl engine as the xlrd engine (being the default) only supports .xls.
# Starting from the 3rd row.
# Skipping rows 4, 5 and 6 since they are empty/not required.
df = pd.read_excel('GL 1000.xlsx', engine='openpyxl', sheet_name='Sales Pg 1', header=2, skiprows=[3,4,5])
df.head()

df.info()

# Convert the Date column to DateTime format instead of object.
df['Date'] = pd.to_datetime(df['Date'])

# The skipfooter function skips x number of final rows - in this case it is the final row.
df_2 = pd.read_excel('GL 1000.xlsx', engine='openpyxl', sheet_name='Sales Pg 2', skipfooter=1, index_col=False)
df_2.tail()

# Transposing the df, then resetting the index and renaming it based on all the columns from the first df.
# When you .reset_index() - this changes the old index currently in place into a column and adds a new sequentially numbered index.
df_2_transposed = df_2.T.reset_index().set_index(df.columns[0:])
# Now transposing it back and resetting the index again.
# By saying drop=True when resetting the index, you are preventing the old index from being added as an extra column!
# The old index needs to be dropped otherwise there would be 2 indexes!
df_2_final = df_2_transposed.T.reset_index(drop=True)
df_2 = df_2_final

# Concatenating the two dataframes - however, the indexes are not sequential.
df_combined = pd.concat([df, df_2])

# Correcting the indexes to be sequential.
df_combined = df_combined.reset_index(drop=True)

# Replacing all the NaN values in the Debit and Credit columns to be 0 instead of NaN.
df_combined['Debit'].fillna(0, inplace=True)


