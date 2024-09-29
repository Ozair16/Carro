import pandas as pd
import sqlite3


#Read the Excel file into a pandas DataFrame
df = pd.read_excel(r'C:\Users\Admin\Desktop\Carro\Data 1 (settlement).xlsx', sheet_name='Sheet')

#Remove the unwanted columns
df_cleaned = df.drop(columns=['Block No','GL Bal','Postage','Misc', 'Repo Exp', 'Storage', 'Legal Exp', 
                              'LOD', 'JPJ Exp','Unnamed: 23', 'Unnamed: 24'])


print(df_cleaned)  
#Export to excel
df_cleaned.to_excel(r'C:\Users\Admin\Desktop\Carro\Hp Settlement.xlsx', index=False)

#Connect to SQLite database
conn = sqlite3.connect('settlements.db')
cursor = conn.cursor()

#Create a table for the cleaned dataset
df_cleaned.to_sql('settlements', conn, if_exists='replace', index=False)

#Verify the data has been inserted correctly
result = pd.read_sql('SELECT * FROM settlements LIMIT 30;', conn)

#Display the result
print(result)

#Commit and close the connection
conn.commit()
conn.close()

