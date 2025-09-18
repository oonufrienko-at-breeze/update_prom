import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel('prom.xlsx')

# Identify duplicate values in the column
duplicates = df[df.duplicated('Код_товара')]

# View the duplicate values
print(duplicates)
