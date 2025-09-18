import pandas as pd
# import re

# Step 1: Open the Excel files
df_prom = pd.read_excel('./data/prom.xlsx')
df_products = pd.read_excel('./data/products.xls')

# Step 2: Create a new DataFrame for updated data
df_updated = df_prom.copy()

# Step 3: Iterate over the rows in 'prom.xlsx'
for _, row_prom in df_prom.iterrows():
    value = row_prom['Название_позиции']
    if pd.notna(value):
        words = value.split()
        if len(words) > 7:
            model = words[-2]
            fourth_word = words[3]
            dn = fourth_word[2:]
            
            # Step 4: Find the corresponding data in 'products.xls'
            matching_rows = df_products.loc[(df_products['name'].str.contains(model)) & (df_products['name'].str.contains(dn))]

            print("Model:", model, "DN", dn)
            print("Matching:", matching_rows)

            # Step 5: Update the corresponding fields in 'prom-updated.xls'
            if not matching_rows.empty:
                # updated_fields = matching_rows[['kod', 'some_other_field']]  # Update with the desired fields from 'products.xls'
                updated_fields = matching_rows[['kod']]
                print("updated_fields:", updated_fields)
                df_updated.loc[_, 'Код_товара'] = updated_fields['kod'].values[0]

                # df_updated.loc[_, ['Код_товара']] = updated_fields.values.flatten()
                # df_updated.loc[_, ['Код_товара', 'Updated Field 2']] = updated_fields.values.flatten()

# Step 6: Save the updated DataFrame to a new Excel file 'prom-updated.xls'
df_updated.to_excel('./output/prom-updated2.xlsx', index=False)
