import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel('./tovar.xls')

# Get the column names
column_names = df.columns

# Print the column names
print(len(column_names))

print(df.columns[len(df.columns) - 3] )

print(df.columns.ravel())
