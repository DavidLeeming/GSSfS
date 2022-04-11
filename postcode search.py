import pandas as pd, numpy as np, sys, openpyxl as px

cwd = (str(sys.argv[0][:-25]))
print(cwd)

try:
    df = pd.read_excel(str(cwd) + "/newsletter.xlsx")
except Exception as Argument:
    print(str(Argument))

try:
    num_range = range(1,42)
    s = 'BS'
    num_list = list(num_range)
    num_list = map(str, num_list) 
    bs = [s + num_range for num_range in num_list]
    bs = '|'.join(bs)

    num_range = range(1,4)
    s = 'BA'
    num_list = list(num_range)
    num_list = map(str, num_list) 
    ba = [s + num_range for num_range in num_list]
    ba = '|'.join(ba)


    df['Trunc'] = df['Postcode'].str[:4]
    conditions = [
        df["Trunc"].str.contains(ba, na=False),
        df["Trunc"].str.contains(bs, na=False)
        ]

    choices = ['BA', 'BS']
    df['Test'] = np.select(conditions, choices, default='NA')

    false = df[(df['Test'] == 'NA')].index 
    df.drop(false , inplace=True)
    df = df.reset_index(drop=True)
    print(df)  
    with pd.ExcelWriter(str(cwd) + "/newsletter.xlsx",engine="openpyxl", mode='a') as writer:
        df.to_excel(writer, sheet_name='BS&BA', index=False)

    wb= px.load_workbook(str(cwd) + '/newsletter.xlsx')
    wb['BS&BA'].auto_filter.ref = wb['BS&BA'].dimensions
    wb.save(str(cwd) + '/newsletter.xlsx')

except Exception as Argument:
    print(str(Argument))