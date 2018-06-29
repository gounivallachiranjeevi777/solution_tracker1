import mammoth
import codecs
from bs4 import BeautifulSoup
import xlrd 
import pandas as pd
import textwrap
import array
import time
import xlrd
import os
import docx2txt
from textblob import TextBlob
from nltk import BlanklineTokenizer
import re

z = os.getcwd()
workbook = xlrd.open_workbook(z+'\\SampleInput.xlsx')
sheets = workbook.sheet_names()
required_data = []
for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data.append((row_values[4]))
required_data2 = []
for sheet_name in sheets:
    sh = workbook.sheet_by_name(sheet_name)
    for rownum in range(sh.nrows):
        row_values = sh.row_values(rownum)
        required_data2.append((row_values[5]))        
required_data1 = list(filter(None, required_data))
z = os.getcwd()
text = docx2txt.process(z+"\\SampleInputDoc1-FAQs.docx")
blob = TextBlob(text)
tokenizer = BlanklineTokenizer()
z = blob.tokenize(tokenizer)
c = '?'
lst = list()
for i in range (0,len(z)):
    x = z[i].find(c)
    if x!= -1:
        lst.append(z[i])

import xlsxwriter

workbook = xlsxwriter.Workbook('SampleOutput.xlsx')
worksheet = workbook.add_worksheet()


# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for problem in (required_data):
    worksheet.write(row, col, problem)
    row += 1

row = 0
col = 1

for solution in (required_data2):
    worksheet.write(row, col, solution)
    row += 1


row = 400
col = 0

# Iterate over the data and write it out row by row.
for i in range (0,len(z)):
    if i%2 == 0:
        worksheet.write(row, col, z[i])
        row += 1

row = 400
col = 1

for i in range (0,len(z)):
    if i%2 != 0:
        worksheet.write(row, col, z[i])
        row += 1

workbook.close()

z = os.getcwd()
f = open(z+'\\SampleInputDoc2-.docx','rb')
b = open('x.html','wb')
document = mammoth.convert_to_html(f)
b.write(document.value.encode('utf8'))
f.close()
b.close()
r=codecs.open("x.html", 'r').read()
soup = BeautifulSoup(r,"lxml")
company_name = soup.find_all('strong')
company_name1 = soup.find_all('h3')

lst = list()
asd = company_name[0]
a = re.sub("<.*?>", "", asd.text)
lst.append(a)
for i in range (0,len(company_name1)):
    x = company_name1[i]
    a = re.sub("<.*?>", "", x.text)
    lst.append(a)

ad = []
for i in range (1,16):
            a = company_name[i]
            a = re.sub("<.*?>", "", a.text)
            ad.append(a)
az = "If you're having problems loading up Windows Explorer and browsing your file system, the problem is almost always a shell extension that shouldn't be installed, or some shell extensions that are conflicting with each other. For example, the shell extensions for Dropbox and TortoiseSVN tend to cause problems when you put your code into your Dropbox folder, causing hanging and generally slow file browsing.Your best bet is to grab a copy of ShellExView and start disabling third-party shell extensions, or uninstalling Windows Explorer plug-ins that you don't actually need. You can also use this tool in combination with ShellMenuView to clean up your messy Explorer context menu."
ad.append(az)

import xlsxwriter

workbook = xlsxwriter.Workbook('SampleOutput1.xlsx')
worksheet = workbook.add_worksheet()


# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for i in range (0,4):
    worksheet.write(row, col, lst[i])
    row += 1

row = 0
col = 1

worksheet.write(row, col, ad[0]+". "+ad[1]+". "+ad[2]+"."+ad[3]+".")
row+=1

worksheet.write(row, col, ad[4]+". "+ad[5]+". "+ad[6]+"."+ad[7]+"."+ad[8]+". "+ad[9]+"."+ad[10]+".")
row+=1

worksheet.write(row, col, ad[11]+". "+ad[12]+". "+ad[13]+"."+ad[14]+".")
row+=1

worksheet.write(row, col, az)
workbook.close()

z = os.getcwd()
f = open(z+'\\SampleInputDoc3-Hardware Problems.docx','rb')
b = open('xy.html','wb')
document = mammoth.convert_to_html(f)
b.write(document.value.encode('utf8'))
f.close()
b.close()

r=codecs.open("xy.html", 'r').read()

soup = BeautifulSoup(r,"lxml")
company_name2 = soup.find_all('ul')
company_name1 = soup.find_all('h4')
company_name = soup.find_all('p')
headings = []
h1 = []
lines = []
for i in range (1,len(company_name2)):
            global aaaa
            aaaa = company_name2[i]
            aaaaa = re.sub("<.*?>", "", aaaa.text)
            lines.append(aaaaa)
for i in range (1,len(company_name)):
            global asa
            asa = company_name[i]
            aa = re.sub("<.*?>", "", asa.text)
            headings.append(aa)
for i in range (1,len(company_name1)):
            global aaa
            aaa = company_name1[i]
            aaaaaa = re.sub("<.*?>", "", aaa.text)
            h1.append(aaaaaa)

ans1 = lines[0]
ques1 = "Unresponsive PC"
ques2 = headings[1]
for i in range (1,10):
    ans2 = lines[i]
ques3 = headings[11]
ans3 = lines[10]
ques4 = headings[12]
ans4 = lines[11]
ques5 = headings[13]
ans5 = lines[12]
ques6 = headings[14]
ans6 = lines[13]
ques7 = headings[15]
ans7 = lines[14]
ques8 = headings[17]
ans8 = lines[15]
ques9 = headings[18]
ans9 = lines[16]
ques10 = headings[19]
ans10 = lines[17]
ques11 = headings[20]
ans11 = lines[18]
ques12 = headings[21]
ans12 = lines[19]
ans12_1 = headings[22]
ans13 = lines[20]
ques14 = headings[23]
ans14_1 = headings[25]
ans14 = lines[21]
ques15 = headings[26]
ans15 = lines[22]
ques16 = headings[27]
ans16 = lines[23]
ques17 = headings[29]
ans17_1 = headings[30]
ans17_2 = lines[24]
ans17_3 = headings[31]
ans17_4 = lines[25]
ques18 = headings[34]
ans18 = headings[35]
ques19 = headings[36]
ans19 = lines[26]
ques20 = headings[37]
ans20 = headings[38]
ques21 = headings[43]
ans21 = headings[44]
ans21_1 = lines[27]
ans21_2 = lines[28]
ques22 = headings[46]
ans22 = lines[29]
ques23 = headings[49]
ans23 = lines[30]
ans23_1 = lines[31]
ques24 = headings[51]
ans24 = lines[32]
ques25 = headings[55]
ans25 = headings[56]
ans25_1 = lines[33]
ans25_2 = headings[57]
ans25_3 = headings[58]
ans25_4 = lines[34]
ans25_5 = headings[59]
ans25_6 = lines[35]
ques26 = headings[60]
ans26 = headings[61]
ans26_1 = lines[36]
ques27 = headings[62]
ans27 = lines[37]
ques28 = headings[63]
ans28 = lines[38]
ques29 = headings[64]
ans29 = lines[39]
ques30 = headings[65]
ans30 = lines[40]
ques31 = headings[66]
ans31 = lines[41]
ques = []
ans = []

ques.extend((ques1, ques2, ques3, ques4, ques5, ques6, ques12, ques14, ques17, ques18, ques19, ques20, ques21, ques22, ques23, ques24, ques25, ques26, ques27, ques28, ques29, ques30, ques31))
ans.extend((ans1, ans2, ans3, ans4, ans5, ans6+" "+ques7+" "+ans7+" "+ques8+" "+ans8+" "+ques9+" "+ans9+" "+ques10+" "+ans10+" "+ques11+" "+ans11, ans12+" "+ans12_1+" "+ans13, ans14_1+" "+ans14+" "+ques15+" "+ans15+" "+ques16+" "+ans16, ans17_1+" "+ans17_2+" "+ans17_3+" "+ans17_4, ans18, ans19, ans20, ans21+" "+ans21_1+" "+ans21_2, ans22, ans23+" "+ans23_1, ans24, ans25+" "+ans25_1+" "+ans25_2+" "+ans25_3+" "+ans25_4+" "+ans25_5+" "+ans25_6, ans26+" "+ans26_1, ans27, ans28, ans29, ans30, ans31))


workbook = xlsxwriter.Workbook('SampleOutput2.xlsx')
worksheet = workbook.add_worksheet()


# Start from the first cell. Rows and columns are zero indexed.
row = 0
col = 0

# Iterate over the data and write it out row by row.
for i in range (0,len(ques)):
    worksheet.write(row, col, ques[i])
    row += 1

row = 0
col = 1

for i in range (0,len(ans)):
    worksheet.write(row, col, ans[i])
    row += 1

workbook.close()

# filenames
excel_names = ["SampleOutput.xlsx", "SampleOutput1.xlsx", "SampleOutput2.xlsx"]

# read them in
excels = [pd.ExcelFile(name) for name in excel_names]

# turn them into dataframes
frames = [x.parse(x.sheet_names[0], header=None,index_col=None) for x in excels]

# concatenate them..
combined = pd.concat(frames)

# write it out
combined.to_excel("Output.xlsx", header=False, index=False)
print('Output file generated')
time.sleep(2)
