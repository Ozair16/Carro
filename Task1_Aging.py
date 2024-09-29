import pandas as pd
import sqlite3
import numpy as np
#Read the Excel file into a pandas DataFrame
df = pd.read_excel(r'C:\Users\Admin\Desktop\Carro\Data 3 (Hp Aging).xlsx', sheet_name='Sheet')

#Remove the unwanted columns
df_cleaned = df.drop(columns=['Occupation', '1st Appr Date', '1st Appr By', 'Notice Date', 'Notice', 'Grading', 'Status'])



#If MOF = 0 then LVT MOF, else MOF and remove LVT MOF
df_cleaned['MOF'] = np.where(df_cleaned['MOF'] == 0, df_cleaned['LVT MOF'], df_cleaned['MOF'])
df_cleaned = df_cleaned.drop(columns=['LVT MOF'])


#To clean RAMCI and get the average and drop the existing RAMCI Average
def calculate_ramci_avg(row):
    columns = ['RAMCI Hirer', 'RAMCI G1', 'RAMCI G2']
    # Filter only the values greater than 0
    valid_values = [row[col] for col in columns if row[col] > 0]
    
    # If no columns have a value greater than 0, return 0 or NaN
    if len(valid_values) == 0:
        return 0
    
    # If only 1 column is greater than 0, return that value
    elif len(valid_values) == 1:
        return valid_values[0]
    
    # If 2 or 3 columns are greater than 0, return their average
    else:
        return sum(valid_values) / len(valid_values)

df_cleaned['RAMCI Avg'] = df_cleaned.apply(calculate_ramci_avg, axis=1)
df_cleaned = df_cleaned.drop(columns=['RAMCI Hirer', 'RAMCI G1', 'RAMCI G2','RAMCI Average'])

print(df_cleaned)  
#Export to excel
df_cleaned.to_excel(r'C:\Users\Admin\Desktop\Carro\Hp Aging.xlsx', index=False)

    
#Connect to SQLite database
conn = sqlite3.connect('aging.db')
cursor = conn.cursor()

#Write the cleaned DataFrame to SQL
df_cleaned.to_sql('aging', conn, if_exists='replace', index=False)

#Verify the data has been inserted
result_df = pd.read_sql('SELECT * FROM aging', conn)
print(result_df)

#Commit and close the connection
conn.commit()
conn.close()

