import pandas as pd

df = pd.read_excel('exampletable.xlsx', sheet_name='Sheet1')
print(df.head())
lst=[]
for index, row in df.iterrows():
    for coldex, col_name in enumerate(df.columns):
        if coldex == 0:
            col0=row[col_name] #this is the first column like '0-10'
            print(col0)
        else:
            if row[col_name]>0:
                for i in range(0,row[col_name]): #we're just looping through the value in the cell X number of times creating rows
                    lst.append([col0, col_name])
                    print([col0, col_name])
df2 = pd.DataFrame(lst, columns=['age', 'participated'])
df2.to_excel('exploded.xlsx', sheet_name='Sheet1', index=False)