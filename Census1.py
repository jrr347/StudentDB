# -*- coding: utf-8 -*-
"""
Created on Mon Jul 11 10:08:06 2016

Explore Nikil's data with Pandas

@author: jrr
"""

import pandas as pd

CENSUS1_PATH = '/Users/jrr/Documents/Stevens/StudentDB/Python/census 1 data file.xlsx'
CENSUS1_WRKSHT = 'census1'
CENSUS3_PATH = '/Users/jrr/Documents/Stevens/StudentDB/Python/census 3 data file.xlsx'
CENSUS3_WRKSHT = 'census3'

def load_wksht(path,sheetName):
    ''' load census1 data as provided by Nikhil 
        path is the file path
        sheetName is the name of the worksheet with the data
        return a pandas dataframe with the data or raise an exception '''
    try:
        return pd.ExcelFile(path).parse(sheetName)
        
    except FileNotFoundError:
        print("Can't load census1 data from: {}".format(path))
 #   except xlrd.biffh.XLRDError:
        print("invalid worksheet name: {}".format(sheetName))
        
        
def map_semester(row):
    ''' map_semester maps semester of form 'YYYY[FSABY] to a sortable value.
        row is a row that includes 'Semester'
    '''
    smap ={ 'F': '1', 'S':'2', 'A':'3','B':'4'}
    if len(row['Semester']) == 5:
        return row['Semester'][0:4] + smap.get(row['Semester'][4],'0')
    else:
        return row['Semester']
        
def map_semester_back(row):
    ''' map_semester maps semester of form 'YYYY[FSABY] to a sortable value.
        row is a row that includes 'Semester'
    '''
    smap ={ '1':'F', '2':'S', '3':'A','4':'B'}
    if len(row['Semester_sort']) == 5:
        return row['Semester_sort'][0:4] + smap.get(row['Semester_sort'][4],'?')
    else:
        return row['Semester_sort']
        

# load the census spreadsheets
census1 = load_wksht(CENSUS1_PATH, CENSUS1_WRKSHT)
census3 = load_wksht(CENSUS3_PATH, CENSUS3_WRKSHT)

'''
# create a unique list of cwids from census1/students
cwids = census1.SID.drop_duplicates()
assert cwids.count() == census1.SID.nunique()
'''
# add a new Semester_sort column to census1 with a sortable version of the Semester
census1['Semester_sort'] = census1.apply(map_semester,axis=1)

# build up a students dataframe with CWID, Name, Email, Entry/firstSemester
students = census1.groupby('SID').agg({'Name':min, 'Semester_sort':min}).reset_index()
assert students.SID.count() == census1.SID.nunique()

# find each student's most recent semester (mrs)
mrs = census1.groupby('SID')['Semester_sort'].max().reset_index()
assert mrs.SID.count() == census1.SID.nunique()

# maintain the max semester_sort from mrs
# students2 has the SID and the most recent semester information
students2 = pd.merge(students, mrs, left_on=['SID'], right_on=['SID'], how='left', suffixes=("_students", ""))
assert students2.SID.nunique() == census1.SID.nunique()

# restore semester to standard format
students2['Semester'] = students2.apply(map_semester_back,axis=1)

# add the SID, Col, Deg, Dept, and Maj1 to students from the most recent semester
scddm = census1[['SID','Col','Deg','Dept', 'Maj1','Semester']]
students3 = pd.merge(students2, scddm, how='left', left_on=['SID','Semester'], right_on=['SID','Semester'])
assert students3.SID.nunique() == census1.SID.nunique()

# rename and drop columns in students3 to "final" state
students3.rename(columns={'SID':'CWID', 'Semester':'MR_Semester'}, inplace=True)
students3.drop(['Semester_sort_students', 'Semester_sort'], axis=1, inplace=True)

# Next, generate grades by merging students and grades
grades = pd.merge(census3, students3, how='left', left_on=['CWID'], right_on=['CWID'])
assert grades.CWID.count() == census3.CWID.count()

outputFile = input('Enter .xlsx file name: ')
writer = pd.ExcelWriter(outputFile, engine='xlsxwriter')
students3.to_excel(writer, sheet_name='Students')
grades.to_excel(writer, sheet_name='Grades')
writer.save()




# create a table/df with relevant information about each student including
# CWID, name, entry, current col, dept, major, 
#students = pd.merge(cwids, census)

'''
ssw = students[students.Maj1 == 'SFEN']
print('Found {} distinct SSW students'.format(ssw['SID'].nunique()))
# add a Semester2 attribute that is sortable from Semester -- this is used to find the most recent semester
ssw['Semester2']= apply(map_semester, axis=1)
'''
'''
Need to be careful doing the merge because each student has a record in students 
for each semester that they have at least one class.
That creates duplicate records in the merge
'''
'''
student_grades = pd.merge(students, grades, left_on='SID', right_on='CWID', how='inner')
x[['Name','Course No','Semester_y' , 'Grade']][ x.SID==10413369]
students.groupby(['Maj1']).count()

# show classes by student ID
print()
'''