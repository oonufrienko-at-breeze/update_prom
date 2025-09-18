
import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel('./tovar.xls')

# Access a specific column by number (e.g., column at index 2)

# for column_index in range(7):
column_index = 5
column = df.iloc[:, column_index]
print("Column ", column)

# Access a specific row by number (e.g., row at index 3)
# row_index = 3
# row = df.iloc[row_index, :]
# print(row)

# # Access a specific cell by row and column numbers (e.g., row 1, column 2)
# cell_value = df.iloc[1, 2]
# print(cell_value)
