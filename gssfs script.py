import datetime, pandas as pd, numpy as np, re, itertools, threading, time, sys
from datetime import date, timedelta

done = 'False'
fail = 'False'
startTime = time.time()

# Method to validate postcodes
def validate_postcode(pc):
    pattern = 'Invalid UK Postcode'
    if len(pc.replace(" ", "")) == 5:
        pattern = re.compile("^[a-zA-Z]{1}[0-9]{2}[a-zA-Z]{2}")
    elif len(pc.replace(" ", "")) == 6:
        pattern = re.compile("^[a-zA-Z]{2}[0-9]{2}[a-zA-Z]{2}")
 #e.g. TW218FF
    elif len(pc.replace(" ", "")) == 7:
        pattern = re.compile("^[a-zA-Z]{2}[0-9]{3}[a-zA-Z]{2}")
    if pattern != 'Invalid UK Postcode':
        if pattern.match(pc):
            return('Valid UK Postcode')
        else:
            return(pattern)

def animate():
    # Function to produce an animation when loading       
    for c in itertools.cycle(['.    ', '..   ', '...  ', '.... ', '...  ', '..   ', '.    ']):
        if done == 'True':
            # Ends script when done == 'True' 
            break
        executionTime = (time.time() - startTime)
        executionTime = round(executionTime, 1)
        sys.stdout.write('\r' + str("Script running! ") + str('Current runtime: ') + str(executionTime) + str('s ') + c) 
        sys.stdout.flush()
        time.sleep(0.15)
        # Time taken between animation cycles is 0.15 seconds
    if fail == 'False':
        sys.stdout.write('\rDone! Script Finished Successfully! ' + str('Runtime: ') + str(executionTime) + str("s"))
        time.sleep(1)
    if fail == 'True':
        sys.stdout.write('\rFailed!!!!!! ' + str('Runtime: ') + str(executionTime) + str("s"))
        time.sleep(1)
    # Once script is finished will print the above message

t = threading.Thread(target=animate)
t.start()

# Memberspace sheet
try:
    df = pd.read_csv("C:/script/download/download.csv")
    # Check for denied users
    df = df[df['Status'] != "denied"]
    # Drop unused cols
    df = df.drop(columns=['Notes', "Plan Status", "Member Plan(s)", "What is your role?", "Please tick if you are a member of:", "Status", " Age of young people taking part: ", "What are you registering as?"])
    # Rename columns
    df.columns = ["First Name", "Last Name", "Email", "Date", "Recruitment method", "Previous participant", "Students", "Organisation", "Postcode", "LA", "GSSfS newsletter", "SEERIH newsletter"]
    # Exclude team members by postcode
    Postcodes_Exclude = ["WN4 9DS", "M16 0HR", "M24 1WH", "CV5 6AL", "M41 6PN"]
    df = df[df.Postcode.isin(Postcodes_Exclude) == False]
    # Remove memberspace import artifacts
    df['Students'] = df['Students'].str.replace('â€“','to')
    # Remove spaces from postcode
    df['Postcode'] = df['Postcode'].str.replace(' ','')
    df['Postcode'] = df['Postcode'].str.upper()
    # Split date and time
    new = df["Date"].str.split(" ", n = 1, expand = True)
    df["Time"] = new[1]
    df["Date"] = new[0]
    df["Time"] = df["Time"].str[:-6]
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

# Student count sheet
try:
    Student_count = df
    # Convert text to numerical values
    conditions = [
        (Student_count['Students'] == "Less than 5"),
        (Student_count['Students'] == "5 – 10"),
        (Student_count['Students'] == "11 – 50"),
        (Student_count['Students'] == "51 – 100"),
        (Student_count['Students'] == "101 – 200"),
        (Student_count['Students'] == "201 – 300"),
        (Student_count['Students'] == "More than 300"),
        ]
    values = [5, 10, 50, 100, 200, 300, 400]
    Student_count['Students'] = np.select(conditions, values)
    Student_count = Student_count.drop(columns=["First Name", "Last Name", "Email", "Recruitment method", "Previous participant", "Organisation", "Postcode", "LA", "GSSfS newsletter", "SEERIH newsletter", "Time"])
    # Number of students per 5 days
    Student_count= Student_count.groupby(pd.Grouper(key='Date', axis=0, freq='5D')).sum()
    Student_count['cum_sum'] = Student_count['Students'].cumsum()
    # Remove time from index
    Student_count.index = Student_count.index.map(lambda x: str(x)[:-9])

    csv_2020 = pd.read_csv("C:/script/2020 number per day.csv")
    csv_2020["Date"] = csv_2020["Year"].astype(str) + csv_2020["Month"].astype(str) + csv_2020["Day"].astype(str)
    csv_2020['Date'] = pd.to_datetime(csv_2020['Date'], format='%Y%B%d') + timedelta(days=365)
    csv_2020 = csv_2020[["Date", "Attendees"]]
    csv_2020.drop(columns=['Attendees'])
    csv_2020= csv_2020.groupby(pd.Grouper(key='Date', axis=0, freq='5D')).sum()
    csv_2020['2020 total'] = csv_2020['Attendees'].cumsum()
    csv_2020.index = csv_2020.index.map(lambda x: str(x)[:-9])
    Student_count = pd.concat([Student_count, csv_2020], axis=1, join="inner")
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

# Target signups

try:
    Target_students = 211898
    Target_signups = 1537


    d0 = date.today()
    d1 = date(2022, 6, 14)
    delta = d1 - d0
    days_left = delta.days
    last_week = d0 - datetime.timedelta(days=7)
    Current_signups = len(df.index)
    df['cum sum'] = df['Students'].cumsum()
    Current_students = df["cum sum"].iloc[-1]
    #weekly_signups = 
    students_per_day = (Target_students - Current_students) / days_left
    d = {'Target Signups': [Target_signups], 'Target Students': [Target_students], 'Current Signups': [Current_signups], 'Current Students': [Current_students], 'Students per day': [students_per_day]}
    target = pd.DataFrame(data=d)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

# List for Manchester
Bury = ['BL0', 'BL8', 'BL9', 'M26', 'M45', 'M25']
Bolton = ['BL1','BL2', 'BL3', 'BL4', 'BL5', 'BL6', 'BL7']
Wigan = ['M29','WA3', 'WN7', 'M46', 'WN2', 'WN4', 'WN3', 'WN5', 'WN6', 'WN1']
Salford = ['M5','M3', 'M7', 'M6', 'M27', 'M28', 'M30', 'M44','M50']
Trafford = ['M16','M32', 'M17', 'M41', 'M33', 'WA15', 'WA14', 'WA13','M31']
Manchester = ['M90','M22', 'M23', 'M20', 'M21', 'M19', 'M16', 'M14','M13', 'M12', 'M18', 'M11', 'M40', 'M9', 'M8', 'M4', 'M3', 'M2', 'M1', 'M15']
Stockport = ['SK7','SK8', 'SK3', 'SK4', 'SK1', 'SK5', 'SK6', 'SK2']
Tameside = ['SK14','M34', 'SK16', 'OL7', 'M43', 'OL6', 'OL5', 'SK15']
Oldham = ['M35','OL8', 'OL4', 'OL3', 'OL2', 'OL1', 'OL9']
Rochdale = ['M24','OL10', 'BL9', 'OL11', 'OL12', 'OL15']
Greater_Manchester = [Bury, Bolton, Wigan, Salford, Trafford, Manchester, Stockport, Tameside, Oldham, Rochdale]
Greater_Manchester = list(itertools.chain.from_iterable(Greater_Manchester))
Greater_Manchester = '|'.join(Greater_Manchester)
# List for East Midlands
Lincolnshire = ['LN1', 'LN2', 'LN3', 'LN4', 'LN5', 'LN6', 'LN7', 'LN8', 'LN9', 'LN10', 'LN11', 'LN12', 'LN13']
Derbyshire = ['DE1', 'DE11','DE13', 'DE12','DE14', 'DE15', 'DE21', 'DE22', 'DE23', 'DE24', 'DE3', 'DE4', 'DE45','DE5','DE55', 'DE56', 'DE6', 'DE65', 'DE7', 'DE72', 'DE73', 'DE74', 'DE75']
Nottinghamshire = ['NG1', 'NG10', 'NG11', 'NG12', 'NG13', 'NG14', 'NG15', 'NG16', 'NG17', 'NG18', 'NG19', 'NG2', 'NG20', 'NG21', 'NG22', 'NG23', 'NG24', 'NG25', 'NG3', 'NG31', 'NG32', 'NG33', 'NG34', 'NG4',
'NG5', 'NG6', 'NG7', 'NG8', 'NG9']
East_Midlands = [Nottinghamshire, Derbyshire, Lincolnshire]
East_Midlands = list(itertools.chain.from_iterable(East_Midlands))
East_Midlands = '|'.join(East_Midlands)
# List for East of England
Cambridgeshire = ['CB1', 'CB10', 'CB11', 'CB2', 'CB21', 'CB22', 'CB23', 'CB24', 'CB25', 'CB3', 'CB4', 'CB5', 'CB6', 'CB7', 'CB8', 'CB9']
Chelmsford = ['CM0', 'CM1', 'CM11', 'CM12', 'CM13', 'CM14', 'CM15', 'CM16', 'CM17', 'CM18', 'CM19', 'CM2', 'CM20', 'CM21', 'CM22', 'CM23', 'CM24', 'CM3', 'CM4', 'CM5', 'CM6', 'CM7', 'CM77', 'CM8', 'CM9']
Colchester = ['CO1', 'CO10', 'CO11', 'CO12', 'CO13', 'CO14', 'CO15', 'CO16', 'CO2', 'CO3', 'CO4', 'CO5', 'CO6', 'CO7', 'CO8', 'CO9']
Ipswich = ['IP1', 'IP10', 'IP11', 'IP12', 'IP13', 'IP14', 'IP15', 'IP16', 'IP17', 'IP18', 'IP19', 'IP2', 'IP20', 'IP21', 'IP22', 'IP23', 'IP24', 'IP25', 'IP26', 'IP27', 'IP28', 'IP29', 'IP3', 'IP30', 'IP31', 'IP32', 
'IP33', 'IP4', 'IP5', 'IP6', 'IP7', 'IP8', 'IP9']
Luton = ['LU1', 'LU2', 'LU3', 'LU4', 'LU5', 'LU6', 'LU7']
Norwich = ['NR1', 'NR10', 'NR11', 'NR12', 'NR13', 'NR14', 'NR15', 'NR16', 'NR17', 'NR18', 'NR19', 'NR2', 'NR20', 'NR21', 'NR22', 'NR23', 'NR24', 'NR25', 'NR26', 'NR27', 'NR28', 'NR29', 'NR3', 'NR30', 'NR31', 'NR32',
'NR33', 'NR34', 'NR35', 'NR4', 'NR5', 'NR6', 'NR8', 'NR9']
Southend_on_sea = ['SS0', 'SS1', 'SS11', 'SS12', 'SS13', 'SS14', 'SS15', 'SS16', 'SS17', 'SS2', 'SS3', 'SS4', 'SS5', 'SS6', 'SS7', 'SS8', 'SS9']
East_of_England = [Cambridgeshire, Chelmsford, Colchester, Ipswich, Luton, Norwich, Southend_on_sea]
East_of_England = list(itertools.chain.from_iterable(East_of_England))
East_of_England = '|'.join(East_of_England)
# List for Guernsey
Guernsey = ['GY1', 'GY2', 'GY3', 'GY4', 'GY5', 'GY6', 'GY7', 'GY8', 'GY9', 'GY10']
Guernsey = '|'.join(Guernsey)
# List for London
London = ['E1', 'E10', 'E11', 'E12', 'E13', 'E14', 'E15', 'E16', 'E17', 'E18', 'E1W', 'E2', 'E3', 'E4', 'E5', 'E6', 'E7', 'E8', 'E9', 'EC1', 'EC1A', 'EC1M', 'EC1N', 'EC1R', 'EC1V', 'EC1Y', 'EC2', 'EC2A', 'EC2M', 
'EC2N', 'EC2R', 'EC2V', 'EC2Y', 'EC3', 'EC3A', 'EC3M', 'EC3N','EC3P', 'EC3R', 'EC3V', 'EC4', 'EC4A', 'EC4M', 'EC4N', 'EC4R', 'EC4V', 'EC4Y', 'N1', 'N10', 'N11', 'N12', 'N13', 'N14', 'N15', 'N16', 'N17', 'N18', 
'N19', 'N2', 'N20', 'N21', 'N22', 'N3', 'N4', 'N5', 'N6', 'N7', 'N8', 'N9', 'NW1', 'NW10', 'NW11', 'NW2', 'NW3', 'NW4', 'NW5', 'NW6', 'NW7', 'NW8', 'NW9', 'SE1', 'SE10', 'SE11', 'SE12', 'SE13', 'SE14', 'SE15', 
'SE16', 'SE17', 'SE18', 'SE19', 'SE2', 'SE20', 'SE21', 'SE22', 'SE23', 'SE24', 'SE25', 'SE26', 'SE27', 'SE28', 'SE3', 'SE4', 'SE5', 'SE6', 'SE7', 'SE8', 'SE9', 'SW1', 'SW10', 'SW11', 'SW12', 
'SW13', 'SW14', 'SW15', 'SW16', 'SW17', 'SW18', 'SW19', 'SW1A', 'SW1E', 'SW1H', 'SW1P', 'SW1V', 'SW1W', 'SW1X', 'SW1Y', 'SW2', 'SW20', 'SW3', 'SW4', 'SW5', 'SW6', 'SW7', 'SW8', 'SW9', 'W1', 'W10', 'W11', 'W12', 
'W13', 'W14', 'W1B', 'W1C', 'W1D', 'W1F', 'W1G', 'W1H', 'W1J', 'W1K', 'W1M', 'W1S', 'W1T', 'W1U', 'W1W', 'W2', 'W3', 'W4', 'W5', 'W6', 'W7', 'W8', 'W9', 'WC1', 'WC1A', 'WC1B', 'WC1E', 'WC1H', 'WC1N', 'WC1R', 
'WC1V', 'WC1X', 'WC2', 'WC2A', 'WC2B', 'WC2E', 'WC2H', 'WC2N', 'WC2R']
London = '|'.join(London)
# List for South West
Bath = ['BA1', 'BA10', 'BA11', 'BA12', 'BA13', 'BA14', 'BA15', 'BA16', 'BA2', 'BA20', 'BA21', 'BA22', 'BA3', 'BA4', 'BA5', 'BA6', 'BA7', 'BA8', 'BA9']
Bristol = ['BS1', 'BS10', 'BS11', 'BS13', 'BS14', 'BS15', 'BS16', 'BS2', 'BS20', 'BS21', 'BS22', 'BS23', 'BS24', 'BS25', 'BS26', 'BS27', 'BS28', 'BS29', 'BS3', 'BS30', 'BS31', 'BS32', 'BS34', 'BS35', 'BS36', 'BS37',
'SB39', 'BS4', 'BS40', 'BS41', 'BS48', 'BS49', 'BS5', 'BS6', 'BS7', 'BS8', 'BS9', 'BS99']
Exeter = ['EX1', 'EX10', 'EX11', 'EX12', 'EX13', 'EX14', 'EX15', 'EX16', 'EX17', 'EX18', 'EX19', 'EX2', 'EX20', 'EX21', 'EX22', 'EX23', 'EX24', 'EX3', 'EX31', 'EX32', 'EX33', 'EX34', 'EX35', 'EX36', 'EX37', 'EX38',
 'EX39', 'EX4', 'EX5', 'EX6', 'EX7', 'EX8', 'EX9']
Plymouth = ['PL1', 'PL10', 'PL11', 'PL12', 'PL13', 'PL14', 'PL15', 'PL16', 'PL17', 'PL18', 'PL19', 'PL2', 'PL20', 'PL21', 'PL22', 'PL23', 'PL24', 'PL25', 'PL26', 'PL27', 'PL28', 'PL29', 'PL3', 'PL30', 'PL31', 'PL32',
'PL32', 'PL33', 'PL34', 'PL35', 'PL4', 'PL5', 'PL6', 'PL7', 'PL8', 'PL9']
Taunton = ['TA1', 'TA10', 'TA11', 'TA12', 'TA13', 'TA14', 'TA15', 'TA16', 'TA17', 'TA18', 'TA19', 'TA2', 'TA20', 'TA21', 'TA22', 'TA23', 'TA24', 'TA3', 'TA4', 'TA5', 'TA6', 'TA7', 'TA8', 'TA9']
Torquay = ['TQ1', 'TQ10', 'TQ11', 'TQ12', 'TQ13', 'TQ14', 'TQ2', 'TQ3', 'TQ4', 'TQ5', 'TQ6', 'TQ7', 'TQ8', 'TQ9']
Truro = ['TR1', 'TR10', 'TR11', 'TR12', 'TR13', 'TR14', 'TR15', 'TR16', 'TR17', 'TR18', 'TR2', 'TR20', 'TR21', 'TR22', 'TR23', 'TR24', 'TR25', 'TR26', 'TR27', 'TR3', 'TR4', 'TR5', 'TR6', 'TR7', 'TR8', 'TR9']
South_West = [Bath, Bristol, Exeter, Plymouth, Taunton, Torquay, Truro]
South_West = list(itertools.chain.from_iterable(South_West))
South_West = '|'.join(South_West)
# List for West Midlands 
Birmingham = ['B1', 'B10', 'B11', 'B12', 'B13', 'B14', 'B15', 'B16', 'B17', 'B18', 'B19', 'B2', 'B20', 'B21', 'B23', 'B24', 'B25', 'B26', 'B27', 'B28', 'B29', 'B3', 'B30', 'B31', 'B32', 'B33', 'B34', 'B35', 
'B36', 'B37', 'B38', 'B4', 'B40', 'B42', 'B43', 'B44', 'B45', 'B46', 'B47', 'B48', 'B49', 'B5', 'B50', 'B6', 'B60', 'B61', 'B62', 'B63', 'B64', 'B65', 'B66', 'B67', 'B68', 'B69', 'B7', 'B70', 'B71', 'B72', 
'B73', 'B74', 'B75', 'B76', 'B77', 'B78', 'B79', 'B8', 'B80', 'B9', 'B90', 'B91', 'B92', 'B93', 'B94', 'B95', 'B96', 'B97', 'B98']
Dudley = ['DY1', 'DY10', 'DY11', 'DY12', 'DY13', 'DY14', 'DY2', 'DY3', 'DY4', 'DY5', 'DY6', 'DY7', 'DY8', 'DY9']
Walsall = ['WS1', 'WS10', 'WS11', 'WS12', 'WS13', 'WS14', 'WS15', 'WS2', 'WS3', 'WS4', 'WS5', 'WS6', 'WS7', 'WS8', 'WS9']
Coventry = ['CV1', 'CV2', 'CV3', 'CV4', 'CV5', 'CV6', 'CV7', 'CV8']
Warwickshire = ['CV9', 'CV10', 'CV11', 'CV12', 'CV13', 'CV21', 'CV22', 'CV23', 'CV31', 'CV32', 'CV33', 'CV34', 'CV35', 'CV36', 'CV37', 'CV47']
West_Midlands = [Birmingham, Dudley, Walsall, Coventry, Warwickshire]
West_Midlands = list(itertools.chain.from_iterable(West_Midlands))
West_Midlands = '|'.join(West_Midlands)
# List for Stoke
Stoke = ['ST1', 'ST10', 'ST11', 'ST12', 'ST13', 'ST14', 'ST15', 'ST16', 'ST17', 'ST18', 'ST19', 'ST2', 'ST20', 'ST21', 'ST3', 'ST4', 'ST5', 'ST6', 'ST7', 'ST8', 'ST9']
Stoke = '|'.join(Stoke)
# List for Yorkshire
Bradford = ['BD1', 'BD10', 'BD11', 'BD12', 'BD13', 'BD14', 'BD15', 'B16', 'BD17', 'BD18', 'BD19', 'BD2', 'BD20', 'BD21', 'BD22', 'BD23', 'BD24', 'BD3', 'BD4', 'BD5', 'BD6', 'BD7', 'BD8', 'BD9']
Halifax = ['HX1', 'HX2', 'HX3', 'HX4', 'HX5', 'HX6', 'HX7']
Wakefield = ['WF1', 'WF10', 'WF11', 'WF12', 'WF13', 'WF14', 'WF15', 'WF16', 'WF17', 'WF2', 'WF3', 'WF4', 'WF5', 'WF6', 'WF7', 'WF8', 'WF9']
Leeds = ['LS1', 'LS10', 'LS11', 'LS12', 'LS13', 'LS14', 'LS15', 'LS16', 'LS17', 'LS18', 'LS19', 'LS2', 'LS20', 'LS21', 'LS22', 'LS23', 'LS24', 'LS25', 'LS26', 'LS27', 'LS28', 'LS29', 'LS3', 'LS4', 'LS5', 
'LS6', 'LS7', 'LS8', 'LS9']
Huddersfield = ['HD1', 'HD2', 'HD3', 'HD4', 'HD5', 'HD6', 'HD7', 'HD8', 'HD9']
Yorkshire_and_Humber = [Bradford, Halifax, Wakefield, Leeds, Huddersfield]
Yorkshire_and_Humber = list(itertools.chain.from_iterable(Yorkshire_and_Humber))
Yorkshire_and_Humber = '|'.join(Yorkshire_and_Humber)
# List for Northern Ireland 
num_range = range(1,95)
s = 'BT'
num_list = list(num_range)
num_list = map(str, num_list) 
Northern_Ireland = [s + num_range for num_range in num_list]
Northern_Ireland = '|'.join(Northern_Ireland)
# List for Scotland 
Aberdeen = ['AB10', 'AB11', 'AB12', 'AB13', 'AB14', 'AB15', 'AB16', 'AB21', 'AB22', 'AB23', 'AB24', 'AB25', 'AB30', 'AB31', 'AB32', 'AB33', 'AB34', 'AB35', 'AB36', 'AB37', 'AB38', 'AB39', 'AB41', 'AB42', 
'AB43', 'AB44', 'AB45', 'AB51', 'AB52', 'AB53', 'AB54', 'AB55', 'AB56', 'AB99']
Dundee = ['DD1', 'DD10', 'DD11', 'DD2', 'DD3', 'DD4', 'DD5', 'DD6', 'DD7', 'DD8', 'DD9']
Dumfries = ['DG1', 'DG10', 'DG11', 'DG12', 'DG13', 'DG14', 'DG16', 'DG2', 'DG3', 'DG4', 'DG5', 'DG6', 'DG7', 'DG8', 'DG9']
num_range = range(1,56)
s = 'EH'
num_list = list(num_range)
num_list = map(str, num_list) 
Edinburgh = [s + num_range for num_range in num_list]
num_range = range(1,22)
s = 'FK'
num_list = list(num_range)
num_list = map(str, num_list) 
Falkirk = [s + num_range for num_range in num_list]
Glasgow = ['G1', 'G11', 'G12', 'G13', 'G14', 'G15', 'G2', 'G20', 'G21', 'G22', 'G23', 'G3', 'G31', 'G32', 'G33', 'G34', 'G4', 'G40', 'G41', 'G42', 'G43', 'G44', 'G45', 'G46', 'G5', 'G51', 'G52', 'G53', 'G60', 
'G61', 'G62', 'G63', 'G64', 'G65', 'G66', 'G67', 'G68', 'G69', 'G71', 'G72', 'G73', 'G74', 'G76', 'G77', 'G78', 'G81', 'G82', 'G83', 'G84']
num_range = range(1,10)
s = 'HS'
num_list = list(num_range)
num_list = map(str, num_list) 
Scottish_Islands = [s + num_range for num_range in num_list]
num_range = range(1,57)
s = 'IV'
num_list = list(num_range)
num_list = map(str, num_list) 
Inverness = [s + num_range for num_range in num_list]
Inverness.insert(1, "IV63")
Inverness = [e for e in Inverness if e not in ('IV50', 'IV37', 'IV38', 'IV39', 'IV29')]
num_range = range(1,31)
s = 'KA'
num_list = list(num_range)
num_list = map(str, num_list) 
Kilmarnock = [s + num_range for num_range in num_list]
num_range = range(1,18)
s = 'KW'
num_list = list(num_range)
num_list = map(str, num_list) 
Orkney = [s + num_range for num_range in num_list]
Orkney = [e for e in Orkney if e not in ('KW4')]
num_range = range(1,18)
s = 'KW'
num_list = list(num_range)
num_list = map(str, num_list) 
Orkney = [s + num_range for num_range in num_list]
Orkney = [e for e in Orkney if e not in ('KW4')]
num_range = range(1,17)
s = 'KY'
num_list = list(num_range)
num_list = map(str, num_list) 
Kirkcaldy = [s + num_range for num_range in num_list]
num_range = range(1,79)
s = 'PA'
num_list = list(num_range)
num_list = map(str, num_list) 
Paisley = [s + num_range for num_range in num_list]
Paisley = [e for e in Paisley if e not in ('PA39', 'PA40', 'PA50')]
num_range = range(1,50)
s = 'PH'
num_list = list(num_range)
num_list = map(str, num_list) 
Perth = [s + num_range for num_range in num_list]
Perth = [e for e in Perth if e not in ('PH27', 'PH28', 'PH29', 'PH37', 'PH45', 'PH46', 'PH47', 'PH48')]
Perth.insert(1, "PH50")
num_range = range(1,16)
s = 'TD'
num_list = list(num_range)
num_list = map(str, num_list) 
Tweed = [s + num_range for num_range in num_list]
Shetland_Isles = ['ZE1', 'ZE2', 'ZE3']
Scotland = [Aberdeen, Dundee, Dumfries, Edinburgh, Falkirk, Glasgow, Scottish_Islands, Inverness, Kilmarnock, Orkney, Kirkcaldy, Paisley,Perth, Tweed]
Scotland = list(itertools.chain.from_iterable(Scotland))
Scotland = '|'.join(Scotland)
# List for Wales
Cardiff = ['CF10', 'CF11', 'CF14', 'CF15', 'CF23', 'CF24', 'CF3', 'CF31', 'CF32', 'CF33', 'CF34', 'CF35', 'CF36', 'CF37', 'CF38', 'CF39', 'CF40', 'CF41', 'CF42', 'CF43', 'CF44', 'CF45', 'CF46', 'CF47', 'CF48', 'CF5', 'CF61', 'CF62', 'CF63', 'CF64', 'CF71', 'CF72', 'CF81', 'CF82', 'CF83']
num_range = range(1,9)
s = 'LD'
num_list = list(num_range)
num_list = map(str, num_list) 
Llandrindod_Wells = [s + num_range for num_range in num_list]
num_range = range(11,79)
s = 'LL'
num_list = list(num_range)
num_list = map(str, num_list) 
Llandudno = [s + num_range for num_range in num_list]
Llandudno = [e for e in Llandudno if e not in ('LL50')]
Newport = ['NP10', 'NP11', 'NP12', 'NP13', 'NP15', 'NP16', 'NP18', 'NP19', 'NP20', 'NP22', 'NP23', 'NP24', 'NP25', 'NP26', 'NP4', 'NP44', 'NP7', 'NP8']
num_range = range(1,74)
s = 'SA'
num_list = list(num_range)
num_list = map(str, num_list) 
Swansea = [s + num_range for num_range in num_list]
Swansea = Swansea = [e for e in Llandudno if e not in ('SA21', 'SA22', 'SA23', 'SA24', 'SA25', 'SA26', 'SA27', 'SA28', 'SA29', 'SA30', 'SA49', 'SA50', 'SA51', 'SA52', 'SA53', 'SA54', 'SA55', 'SA56', 'SA57', 'SA58', 'SA59', 'SA60')]
Wales = [Cardiff, Llandrindod_Wells, Llandudno, Newport, Swansea]
Wales = list(itertools.chain.from_iterable(Wales))
Wales = '|'.join(Wales)

#try:
df['Trunc'] = df['Postcode'].str[:4]
conditions = [
    df["Trunc"].str.contains(London, na=False),
    df["Trunc"].str.contains(Guernsey, na=False),
    df["Trunc"].str.contains(Greater_Manchester, na=False),
    df["Trunc"].str.contains(Yorkshire_and_Humber, na=False),
    df["Trunc"].str.contains(Northern_Ireland, na=False),
    df["Trunc"].str.contains(Scotland, na=False),
    df["Trunc"].str.contains(Wales, na=False),
    df["Trunc"].str.contains(East_Midlands, na=False),
    df["Trunc"].str.contains(East_of_England, na=False),
    df["Trunc"].str.contains(West_Midlands, na=False),
    df["Trunc"].str.contains(South_West, na=False),
    df["Trunc"].str.contains(Stoke, na=False),
]

choices = ['London', 'Guernsey', 'Manchester','Yorkshire', 'Northern Ireland', 'Scotland', 'Wales', 
'East Midlands', 'East of England', 'West Midlands', 'South West', 'Stoke']
df['Region'] = np.select(conditions, choices, default='NA')

try:
    fields = ['pcd7', 'ladcd', "lsoa11cd"]
    postcode_csv = pd.read_csv("C:/script/postcode.csv", usecols=fields)
    postcode_csv.columns = ["Postcode", "lsoa", "ladcd"]
    postcode_csv['Postcode'] = postcode_csv['Postcode'].str.replace(' ','')

    fields = ['Postcode', 'MPM rating']
    MPM_csv = pd.read_excel("C:/script/mpm.xlsx", usecols=fields)
    MPM_csv['Postcode'] = MPM_csv['Postcode'].str.replace(' ','')

    fields = ["LSOA code (2011)", "IDACI"]
    Idaci = pd.read_excel("C:/script/idaci.xlsx", usecols=fields)
    Idaci.columns = ["lsoa", "IDACI"]

    df = pd.merge(left=df, right=postcode_csv, how='left', left_on='Postcode', right_on='Postcode')
    df = pd.merge(left=df, right=MPM_csv, how='left', left_on='Postcode', right_on='Postcode')
    df = pd.merge(left=df, right=Idaci, how='left', left_on='lsoa', right_on='lsoa')

    MPM_df = df['MPM rating'].str.extract('(.{1,1})' * 2)
    MPM_df.columns = ["Deprivation rating", "Education rating"]
    MPM_df["Deprivation rating"] = MPM_df["Deprivation rating"].apply(pd.to_numeric)
    df = pd.concat([df, MPM_df], axis=1)
    df = df.drop(columns=['MPM rating'])

    conditions = [
        (df['Deprivation rating'] >= 2),
        (df['Deprivation rating'] <= 3)
        ]
    values = ["1/2", "3+"]
    df['MPM distribution'] = np.select(conditions, values)
    conditions = [
        (df['IDACI'] <= 4),
        (df['IDACI'] >= 5)
        ]
    values = ["lower 40%", "upper 60 %"]
    df['IDACI %'] = np.select(conditions, values)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

# England recruited
try:
    fields = ["Establishment name", "Postcode", "Phase of education", "Website address", "Telephone number"]
    Eng_school_list = pd.read_excel("C:/script/Eng_Schools.xlsx", 'Open', usecols=fields, converters={'Telephone number':str})
    Eng_school_list['Postcode'] = Eng_school_list['Postcode'].str.replace(' ','')
    Eng_school_list['Trunc'] = Eng_school_list['Postcode'].str[:4]
    conditions = [
        Eng_school_list["Trunc"].str.contains(London, na=False),
        Eng_school_list["Trunc"].str.contains(Guernsey, na=False),
        Eng_school_list["Trunc"].str.contains(Greater_Manchester, na=False),
        Eng_school_list["Trunc"].str.contains(Yorkshire_and_Humber, na=False),
        Eng_school_list["Trunc"].str.contains(East_Midlands, na=False),
        Eng_school_list["Trunc"].str.contains(East_of_England, na=False),
        Eng_school_list["Trunc"].str.contains(West_Midlands, na=False),
        Eng_school_list["Trunc"].str.contains(South_West, na=False),
        Eng_school_list["Trunc"].str.contains(Stoke, na=False),
    ]

    choices = ['London', 'Guernsey', 'Manchester','Yorkshire', 'East Midlands',
    'East of England', 'West Midlands', 'South West', 'Stoke']
    Eng_school_list['Region'] = np.select(conditions, choices, default='NA')
    Eng_school_list = Eng_school_list.drop(columns=['Trunc'])
    Eng_schools_not_recruited = Eng_school_list
    columns = ['Postcode']
    dict = df['Postcode']
    Recruited = pd.DataFrame(dict, columns=columns)
    Eng_schools_not_recruited = pd.concat([Eng_schools_not_recruited,Recruited], join='outer')
    Eng_schools_not_recruited = pd.concat([Eng_schools_not_recruited,Recruited], join='outer')
    Eng_schools_not_recruited = Eng_schools_not_recruited.drop_duplicates(subset=['Postcode'], keep=False)
    Eng_school_list = Eng_school_list.dropna(subset=['Postcode'])
except Exception as Argument:
    fail = 'True'
    print(str(Argument))

# Scotland recruited 
try:
    fields = ["School Name", "Post Code", "Email", "Phone Number", "Website Address", "Primary Department", "Secondary Department"]
    Scot_school_list = pd.read_excel("C:/script/Scot_Schools.xlsx", 'Open Schools', usecols=fields, skiprows= 5)
    Scot_school_list['Postcode'] = Scot_school_list['Post Code'].str.replace(' ','')
    Scot_school_list = Scot_school_list.drop(columns=['Post Code'])
    Scot_schools_not_recruited = Scot_school_list 
    Scot_schools_not_recruited = pd.concat([Scot_schools_not_recruited,Recruited], join='outer')
    Scot_schools_not_recruited = pd.concat([Scot_schools_not_recruited,Recruited], join='outer')
    Scot_schools_not_recruited = Scot_schools_not_recruited.drop_duplicates(subset=['Postcode'], keep=False)
    Scot_school_list = Scot_school_list.dropna(subset=['Postcode'])
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

# Wales recruited
try:
    fields = ["School Name", "Postcode", "Phone Number"]
    Wales_school_list = pd.read_excel("C:/script/Wales_Schools.xlsx", usecols=fields, converters={'Phone Number':str})
    Wales_school_list['Postcode'] = Wales_school_list['Postcode'].str.replace(' ','')
    Wales_schools_not_recruited = Wales_school_list 
    Wales_schools_not_recruited = pd.concat([Wales_schools_not_recruited,Recruited], join='outer')
    Wales_schools_not_recruited = pd.concat([Wales_schools_not_recruited,Recruited], join='outer')
    Wales_schools_not_recruited = Wales_schools_not_recruited.drop_duplicates(subset=['Postcode'], keep=False)
    Wales_school_list = Wales_school_list.dropna(subset=['Postcode'])
except Exception as Argument:
    print(str(Argument))
    fail = 'True'
    
#NI Recruited
try:
    fields = ["Institution_Name", "Postcode", "Telephone", "Email", "Institution_Type"]
    NI_school_list = pd.read_excel("C:/script/NI_Schools.xlsx", usecols=fields, converters={'Telephone':str})
    NI_school_list['Postcode'] = NI_school_list['Postcode'].str.replace(' ','')
    NI_schools_not_recruited = NI_school_list 
    NI_schools_not_recruited = pd.concat([NI_schools_not_recruited,Recruited], join='outer')
    NI_schools_not_recruited = pd.concat([NI_schools_not_recruited,Recruited], join='outer')
    NI_schools_not_recruited = NI_schools_not_recruited.drop_duplicates(subset=['Postcode'], keep=False)
    NI_school_list = NI_school_list.dropna(subset=['Postcode'])
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    column_names = ["First Name", "Last Name", "Email", "Date", "Time", "Recruitment method", "Previous participant", "Students", "cum sum", "Organisation", "Postcode", "Region", "LA", "GSSfS newsletter", "SEERIH newsletter", "Deprivation rating", 'MPM distribution', "Education rating", "IDACI", "IDACI %", "lsoa", "ladcd"]
    df = df.reindex(columns=column_names)
    df['Date'] = pd.to_datetime(df['Date']).dt.date

    df.fillna(str("NA"))
except Exception as Argument:
    print(str(Argument))
    fail = 'True'

try:
    df.to_excel("C:/script/output.xlsx", sheet_name='Memberspace', index=False)   
    with pd.ExcelWriter("C:/script/output.xlsx",engine="openpyxl", mode='a') as writer:
        target.to_excel(writer, sheet_name='Target Signups', index=False)
        Student_count.to_excel(writer, sheet_name='Student Count')
        Eng_school_list.to_excel(writer, sheet_name='Eng List', index=False)
        Eng_schools_not_recruited.to_excel(writer, sheet_name='Eng Not recruited', index=False)
        Scot_school_list.to_excel(writer, sheet_name='Scot List', index=False)
        Scot_schools_not_recruited.to_excel(writer, sheet_name='Scot Not recruited', index=False)
        Wales_school_list.to_excel(writer, sheet_name='Wales List', index=False)
        Wales_schools_not_recruited.to_excel(writer, sheet_name='Wales Not recruited', index=False)
        NI_school_list.to_excel(writer, sheet_name='NI List', index=False)
        NI_schools_not_recruited.to_excel(writer, sheet_name='NI Not recruited', index=False)
except Exception as Argument:
    print(str(Argument))
    fail = 'True'


done = 'True'