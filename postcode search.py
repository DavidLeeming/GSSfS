import pandas as pd, numpy as np, sys
cwd = (str(sys.argv[0][:-25]))
print(cwd)

try:
    fields = ["Organisation", "Postcode"]
    df = pd.read_excel(str(cwd) + "/output.xlsx", sheet_name='Memberspace', usecols=fields)
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


    

except Exception as Argument:
    print(str(Argument))