import datetime, pandas as pd, numpy as np
from datetime import date, timedelta

# Memberspace sheet

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

# Student count sheet

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
#csv_2020["Date"] = csv_2020['Date'] + pd.offsets.DateOffset(years=1)
Student_count = pd.concat([Student_count, csv_2020], axis=1, join="inner")
# Target signups

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


fields = ['pcd7', 'ladcd', "lsoa11cd"]
postcode_csv = pd.read_csv("C:/script/postcode.csv", usecols=fields)
postcode_csv.columns = ["Postcode", "lsoa", "ladcd"]
postcode_csv['Postcode'] = postcode_csv['Postcode'].str.replace(' ','')
fields = ['LAD21CD', 'ITL121NM']
region_csv = pd.read_excel("C:/script/regions.xlsx", usecols=fields)
region_csv.columns = ["ladcd", "Region"]
regional = (pd.merge(postcode_csv, region_csv, on='ladcd'))

fields = ['Postcode', 'MPM rating']
MPM_csv = pd.read_excel("C:/script/mpm.xlsx", usecols=fields)
MPM_csv['Postcode'] = MPM_csv['Postcode'].str.replace(' ','')

fields = ["LSOA code (2011)", "IDACI"]
Idaci = pd.read_excel("C:/script/idaci.xlsx", usecols=fields)
Idaci.columns = ["lsoa", "IDACI"]


df = pd.merge(left=df, right=regional, how='left', left_on='Postcode', right_on='Postcode')
df = pd.merge(left=df, right=MPM_csv, how='left', left_on='Postcode', right_on='Postcode')
df = pd.merge(left=df, right=Idaci, how='left', left_on='lsoa', right_on='lsoa')

MPM_df = df['MPM rating'].str.extract('(.{1,1})' * 2)
MPM_df.columns = ["Deprivation rating", "Education rating"]
#MPM_df["Deprivation rating"] = pd.to_numeric(MPM_df["Deprivation rating"])
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

# England recruited
fields = ["Establishment name", "Postcode", "Phase of education", "Website address", "Telephone number"]
Eng_school_list = pd.read_excel("C:/script/Eng_Schools.xlsx", 'Open', usecols=fields, converters={'Telephone number':str})
Eng_school_list['Postcode'] = Eng_school_list['Postcode'].str.replace(' ','')
Eng_schools_not_recruited = Eng_school_list
columns = ['Postcode']
dict = df['Postcode'] 
Recruited = pd.DataFrame(dict, columns=columns)
Eng_schools_not_recruited = pd.concat([Eng_schools_not_recruited,Recruited], join='outer')
Eng_schools_not_recruited = pd.concat([Eng_schools_not_recruited,Recruited], join='outer')
Eng_schools_not_recruited = Eng_schools_not_recruited.drop_duplicates(subset=['Postcode'], keep=False)
Eng_school_list = Eng_school_list.dropna(subset=['Postcode'])

# Scotland recruited 
fields = ["School Name", "Post Code", "Phase of education", "Email", "Phone Number", "Website Address", "Primary Department", "Secondary Department"]
Scot_school_list = pd.read_excel("C:/script/Scot_Schools.xlsx", 'Open Schools', usecols=fields)
Scot_school_list['Postcode'] = Scot_school_list['Post code'].str.replace(' ','')
Scot_school_list = Scot_school_list.drop(columns=['Post code'])
Scot_schools_not_recruited = Scot_school_list 
Scot_schools_not_recruited = pd.concat([Scot_schools_not_recruited,Recruited], join='outer')
Scot_schools_not_recruited = pd.concat([Scot_schools_not_recruited,Recruited], join='outer')
Scot_schools_not_recruited = Scot_schools_not_recruited.drop_duplicates(subset=['Postcode'], keep=False)
Scot_school_list = Scot_school_list.dropna(subset=['Postcode'])





column_names = ["First Name", "Last Name", "Email", "Date", "Time", "Recruitment method", "Previous participant", "Students", "cum sum", "Organisation", "Postcode", "Region", "LA", "GSSfS newsletter", "SEERIH newsletter", "Deprivation rating", 'MPM distribution', "Education rating", "IDACI", "IDACI %", "lsoa", "ladcd"]
df = df.reindex(columns=column_names)
df['Date'] = pd.to_datetime(df['Date']).dt.date
print(df)
df.fillna(str("NA"))
df.to_excel("C:/script/output.xlsx", sheet_name='Memberspace', index=False)   
with pd.ExcelWriter("C:/script/output.xlsx",engine="openpyxl", mode='a') as writer:
    target.to_excel(writer, sheet_name='Target Signups', index=False)
    Student_count.to_excel(writer, sheet_name='Student Count', index=False)
    Eng_school_list.to_excel(writer, sheet_name='Eng List', index=False)
    Eng_schools_not_recruited.to_excel(writer, sheet_name='Eng Not recruited', index=False)