import xml.etree.ElementTree as ET
import xlsxwriter
from pprint import pprint

filename = "Fortify-Developer-Workbook-Report.xml"
fileResult = "Converted-Report.xlsx"

events=("start", "end")
path = []

resultSection = False

data = {
  "finding": "",
  "priority": "",
  "status": "",
  "count": "",
  "summary": "",
  "filepath": "",
  "loc": 0,
  "component": "",
  "dev": "",
  "sec": ""
}

def writeExcel(data, row):
    #print(data['count'])
    #print(data['finding'])
    #print(data['summary'])
    #print(data['priority'])
    #print(data['filepath'])
    #print(data['loc'])
    #print(data['component'])

    for col, item in enumerate(data):
        worksheet.write(row, col, data[item])

row = 1

workbook = xlsxwriter.Workbook(fileResult)
worksheet = workbook.add_worksheet()

#Write header
for col, item in enumerate(data):
    worksheet.write(0, col, item.capitalize())

#Write additional header
worksheet.write(0, 8, "Developer Remark")
worksheet.write(0, 9, "Security Remark")

parser = ET.XMLParser(encoding='utf-8')

for event, elem in ET.iterparse(filename, events=events, parser=parser):
    if event == 'start':
        #Result section
        if not resultSection and elem.tag == 'Title' and elem.text == 'Results Outline':
            resultSection = True

        if resultSection:            
            #Get number of issues
            if elem.tag == 'GroupingSection':
                data['count'] = elem.attrib['count']
                #print(data['count'])

            #Get issue title
            if elem.tag == 'groupTitle':
                data['finding'] = elem.text
                #print(data['finding'])

            #Get analysis
            if elem.tag == 'Value' and elem.text == 'Not an Issue':
                data['status'] = 'False Positive'
                #pprint(data['status'])

            #Get comment
            if elem.tag == 'Comment':
                data['sec'] = elem.text
                #pprint(data['sec'])
            
            #Get abstract
            if elem.tag == 'Abstract':
                data['summary'] = elem.text
                #pprint(data['summary'])

            #Get priority
            if elem.tag == 'Friority':
                data['priority'] = elem.text
                #print(data['priority'])

            #Get filepath
            if elem.tag == 'FilePath':
                data['filepath'] = elem.text
                #if data['finding'] == 'Insecure Randomness':
                    #pprint(data['filepath'])
                    #pprint(ET.tostring(elem).decode())
                    
            #get line of code
            if elem.tag == 'LineStart':
                data['loc'] = elem.text
                #print(data['loc'])

            #get affected component
            if elem.tag == 'TargetFunction':
                data['component'] = elem.text
                #print(data['component'])

                print('Processing data ' + str(row))
                writeExcel(data, row)
                row += 1

                #Clear fields value
                data['summary'] = ''
                data['priority'] = ''
                data['filepath'] = ''
                data['loc'] = 0
                data['component'] = ''
                data['sec'] = ''
                data['status'] = ''
                
                #break

    elif event == 'end':
        elem.clear()
        #if row>=20: break #remove, testing only
            
workbook.close()
