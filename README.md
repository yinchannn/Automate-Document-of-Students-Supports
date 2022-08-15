# Automating Documentation of Supporting Measures for Students with Diverse Needs
The following Python codes aim to generate a record sheet (.docx file) of the supporting measures conducted to each individual student with special education need(s) in the specified academic year from a template file.
<br/><br/>
A record sheet generated should contain the following items:
- a filename in the format of "Class Student Name.docx"
- student name stated in appropriate place
- class stated in appropriate place
- learning difficulties stated in appropriate place
- the one or two elective subject(s) taken by the student stated in appropriate place
- ticks indicating the supporting measures conducted in specified subjects
- student's scores in each subject in 1st term exam
- student's average score in 1st term exam
- student's form position in 1st term exam
<br/><br/>

The necessary information are stored in different excel files:
- Students_Info.xlsx


<br/>
Note that all raw data in this demonstration are anoymized and modified in order to protect students and the school's privacy.
<br/>

Install python-docx and import the necessary packages to Python.

```python
!pip install python-docx
from docx import Document
import pandas as pd
import numpy as np
```
<br/>
Read students' information from the excel file. Convert it into a dataframe and clean it.

```python
#Getting students' info
info = pd.read_excel('Students_Info.xlsx')
info = pd.DataFrame(info)

#Clean info dataframe
info.rename(columns={'SEN\nCode':'SEN Code'}, inplace=True)
info['SEN Code'] = info['SEN Code'].astype(int)

#Set index
info.set_index(['SEN Code'], inplace=True)
```
<img src="Students_Info.png" width="500">
<br/>
Read the template file. Convert all tables into one dataframe and clean it.

```python
#Read the template file
doc = Document('template_s5.docx')

#Convert tables into one dataframe
table_list = []
for table in doc.tables:
    data = [[cell.text for cell in row.cells] for row in table.rows]
    table_list.append(pd.DataFrame(data))
df = pd.concat(table_list)

#Set header
df.columns = ['Item', 'Supporting Measure', 'Chi', 'Eng', 'Mat', 'LS', 'X1', 'X2', 'RE', 'PE', 'Avg', 'Position', 'Position']

#Clean the dataframe
df['Item'] = df['Item'].str.strip()
df['Supporting Measure'] = df['Supporting Measure'].str.strip()

#Set index
df.set_index('Supporting Measure', inplace=True)
```
<img src="Template.png" width="500">
<br/>
