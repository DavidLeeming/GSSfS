try:
    Bury = ['BL0', 'BL8', 'BL9', 'M26', 'M45', 'M25']
    Bury_df = df
    Bury_df['Trunc'] = Bury_df['Postcode'].str[:3]
    pattern = '|'.join(Bury)
    Bury_df = Bury_df[Bury_df.Trunc.str.contains(pattern) == True]
    Bury_df = Bury_df.drop(columns=['Trunc'])
    Man_school_list = Eng_school_list
    Man_school_list['Trunc'] = Man_school_list['Postcode'].str[:3]
    Man_school_list = Man_school_list[Man_school_list.Trunc.str.contains(pattern) == True]
    Bury_count = len(Man_school_list.index)
    Bury_current_count = len(Bury_df.index)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    Bolton = ['BL1','BL2', 'BL3', 'BL4', 'BL5', 'BL6', 'BL7']
    Bolton_df = df
    Bolton_df['Trunc'] = Bolton_df['Postcode'].str[:3]
    pattern = '|'.join(Bolton)
    Bolton_df = Bolton_df[Bolton_df.Trunc.str.contains(pattern) == True]
    Bolton_df = Bolton_df.drop(columns=['Trunc'])
    Man_school_list = Eng_school_list
    Man_school_list['Trunc'] = Man_school_list['Postcode'].str[:3]
    Man_school_list = Man_school_list[Man_school_list.Trunc.str.contains(pattern) == True]
    Bolton_count = len(Man_school_list.index)
    Bolton_current_count = len(Bolton_df.index)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    Wigan = ['M29','WA3', 'WN7', 'M46', 'WN2', 'WN4', 'WN3', 'WN5', 'WN6', 'WN1']
    Wigan_df = df
    Wigan_df['Trunc'] = Wigan_df['Postcode'].str[:3]
    pattern = '|'.join(Wigan)
    Wigan_df = Wigan_df[Wigan_df.Trunc.str.contains(pattern) == True]
    Wigan_df = Wigan_df.drop(columns=['Trunc'])
    Man_school_list = Eng_school_list
    Man_school_list['Trunc'] = Man_school_list['Postcode'].str[:3]
    Man_school_list = Man_school_list[Man_school_list.Trunc.str.contains(pattern) == True]
    Wigan_count = len(Man_school_list.index)
    Wigan_current_count = len(Wigan_df.index)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    Salford = ['M5','M3', 'M7', 'M6', 'M27', 'M28', 'M30', 'M44','M50']
    Salford_df = df
    Salford_df['Trunc'] = Salford_df['Postcode'].str[:3]
    pattern = '|'.join(Salford)
    Salford_df = Salford_df[Salford_df.Trunc.str.contains(pattern) == True]
    Salford_df = Salford_df.drop(columns=['Trunc'])
    Man_school_list = Eng_school_list
    Man_school_list['Trunc'] = Man_school_list['Postcode'].str[:3]
    Man_school_list = Man_school_list[Man_school_list.Trunc.str.contains(pattern) == True]
    Salford_count = len(Man_school_list.index)
    Salford_current_count = len(Salford_df.index)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    Trafford = ['M16','M32', 'M17', 'M41', 'M33', 'WA15', 'WA14', 'WA13','M31']
    Trafford_df = df
    Trafford_df['Trunc'] = Trafford_df['Postcode'].str[:3]
    pattern = '|'.join(Trafford)
    Trafford_df = Trafford_df[Trafford_df.Trunc.str.contains(pattern) == True]
    Trafford_df = Trafford_df.drop(columns=['Trunc'])
    Man_school_list = Eng_school_list
    Man_school_list['Trunc'] = Man_school_list['Postcode'].str[:3]
    Man_school_list = Man_school_list[Man_school_list.Trunc.str.contains(pattern) == True]
    Trafford_count = len(Man_school_list.index)
    Trafford_current_count = len(Trafford_df.index)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    Manchester = ['M90','M22', 'M23', 'M20', 'M21', 'M19', 'M16', 'M14','M13', 'M12', 'M18', 'M11', 'M40', 'M9', 'M8', 'M4', 'M3', 'M2', 'M1', 'M15']
    Manchester_df = df
    Manchester_df['Trunc'] = Manchester_df['Postcode'].str[:3]
    pattern = '|'.join(Manchester)
    Manchester_df = Manchester_df[Manchester_df.Trunc.str.contains(pattern) == True]
    Manchester_df = Manchester_df.drop(columns=['Trunc'])
    Man_school_list = Eng_school_list
    Man_school_list['Trunc'] = Man_school_list['Postcode'].str[:3]
    Man_school_list = Man_school_list[Man_school_list.Trunc.str.contains(pattern) == True]
    Manchester_count = len(Man_school_list.index)
    Manchester_current_count = len(Manchester_df.index)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    Stockport = ['SK7','SK8', 'SK3', 'SK4', 'SK1', 'SK5', 'SK6', 'SK2']
    Stockport_df = df
    Stockport_df['Trunc'] = Stockport_df['Postcode'].str[:3]
    pattern = '|'.join(Stockport)
    Stockport_df = Stockport_df[Stockport_df.Trunc.str.contains(pattern) == True]
    Stockport_df = Stockport_df.drop(columns=['Trunc'])
    Man_school_list = Eng_school_list
    Man_school_list['Trunc'] = Man_school_list['Postcode'].str[:3]
    Man_school_list = Man_school_list[Man_school_list.Trunc.str.contains(pattern) == True]
    Stockport_count = len(Man_school_list.index)
    Stockport_current_count = len(Stockport_df.index)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    Tameside = ['SK14','M34', 'SK16', 'OL7', 'M43', 'OL6', 'OL5', 'SK15']
    Tameside_df = df
    Tameside_df['Trunc'] = Tameside_df['Postcode'].str[:3]
    pattern = '|'.join(Tameside)
    Tameside_df = Tameside_df[Tameside_df.Trunc.str.contains(pattern) == True]
    Tameside_df = Tameside_df.drop(columns=['Trunc'])
    Man_school_list = Eng_school_list
    Man_school_list['Trunc'] = Man_school_list['Postcode'].str[:3]
    Man_school_list = Man_school_list[Man_school_list.Trunc.str.contains(pattern) == True]
    Tameside_count = len(Man_school_list.index)
    Tameside_current_count = len(Tameside_df.index)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    Oldham = ['M35','OL8', 'OL4', 'OL3', 'OL2', 'OL1', 'OL9']
    Oldham_df = df
    Oldham_df['Trunc'] = Oldham_df['Postcode'].str[:3]
    pattern = '|'.join(Oldham)
    Oldham_df = Oldham_df[Oldham_df.Trunc.str.contains(pattern) == True]
    Oldham_df = Oldham_df.drop(columns=['Trunc'])
    Man_school_list = Eng_school_list
    Man_school_list['Trunc'] = Man_school_list['Postcode'].str[:3]
    Man_school_list = Man_school_list[Man_school_list.Trunc.str.contains(pattern) == True]
    Oldham_count = len(Man_school_list.index)
    Oldham_current_count = len(Oldham_df.index)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    Rochdale = ['M24','OL10', 'BL9', 'OL11', 'OL12', 'OL15']
    Rochdale_df = df
    Rochdale_df['Trunc'] = Rochdale_df['Postcode'].str[:3]
    pattern = '|'.join(Rochdale)
    Rochdale_df = Rochdale_df[Rochdale_df.Trunc.str.contains(pattern) == True]
    Rochdale_df = Rochdale_df.drop(columns=['Trunc'])
    Man_school_list = Eng_school_list
    Man_school_list['Trunc'] = Man_school_list['Postcode'].str[:3]
    Man_school_list = Man_school_list[Man_school_list.Trunc.str.contains(pattern) == True]
    Rochdale_count = len(Man_school_list.index)
    Rochdale_current_count = len(Rochdale_df.index)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    Bury_percent = Bury_current_count/Bury_count
    Bury_percent = Bury_percent * 100
    Bolton_percent = Bolton_current_count/Bolton_count
    Bolton_percent = Bolton_percent * 100
    Wigan_percent = Wigan_current_count/Wigan_count
    Wigan_percent = Wigan_percent * 100
    Salford_percent = Salford_current_count/Salford_count
    Salford_percent = Salford_percent * 100
    Trafford_percent = Trafford_current_count / Trafford_count
    Trafford_percent = Trafford_percent * 100
    Manchester_percent = Manchester_current_count / Manchester_count
    Manchester_percent = Manchester_percent * 100 
    Stockport_percent = Stockport_current_count / Stockport_count
    Stockport_percent = Stockport_percent * 100
    Tameside_percent = Tameside_current_count / Tameside_count
    Tameside_percent = Tameside_percent * 100
    Oldham_percent = Oldham_current_count / Oldham_count
    Oldham_percent = Oldham_percent * 100
    Rochdale_percent = Rochdale_current_count / Rochdale_count
    Rochdale_percent = Rochdale_percent * 100
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    d = {'Local Authority': ['Bury', 'Bolton', 'Wigan', 'Salford', 'Trafford', 'Manchester', 'Stockport', 'Tameside', 'Oldham', 'Rochdale'], 'Current count': [Bury_current_count, Bolton_current_count, Wigan_current_count, Salford_current_count, Trafford_current_count, Manchester_current_count, Stockport_current_count, Tameside_current_count, Oldham_current_count, Rochdale_current_count], 'Total schools': [Bury_count, Bolton_count, Wigan_count, Salford_count, Trafford_count, Manchester_count, Stockport_count, Tameside_count, Oldham_count, Rochdale_count], '% recruited': [Bury_percent, Bolton_percent, Wigan_percent, Salford_percent, Trafford_percent, Manchester_percent, Stockport_percent, Tameside_percent, Oldham_percent, Rochdale_percent]}
    Totals = pd.DataFrame(data=d)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'


try:
    Totals.to_excel("C:/script/GM areas.xlsx", sheet_name='Total', index=False)
    with pd.ExcelWriter("C:/script/GM areas.xlsx",engine="openpyxl", mode='a') as writer:
        Bury_df.to_excel(writer, sheet_name='Bury', index=False)
        Bolton_df.to_excel(writer, sheet_name='Bolton', index=False)
        Wigan_df.to_excel(writer, sheet_name='Wigan', index=False)
        Salford_df.to_excel(writer, sheet_name='Salford', index=False)
        Trafford_df.to_excel(writer, sheet_name='Trafford', index=False)
        Manchester_df.to_excel(writer, sheet_name='Manchester', index=False)
        Stockport_df.to_excel(writer, sheet_name='Stockport', index=False)
        Tameside_df.to_excel(writer, sheet_name='Tameside', index=False)
        Oldham_df.to_excel(writer, sheet_name='Oldham', index=False)
        Rochdale_df.to_excel(writer, sheet_name='Rochdale', index=False)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'