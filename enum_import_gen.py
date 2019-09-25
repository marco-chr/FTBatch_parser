# Convert enumeration sets contained in BES export .axml file to XLS Import file for NEOPLM
# Release 0.0 -  30 Oct 2018
#
# Usage python enum_import_gen.py filename

import xlsxwriter
import xml.etree.ElementTree as ET
import sys

print("Opening axml file")
tree = ET.parse(sys.argv[1])
root = tree.getroot()

print("Creating workbook")
# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook(sys.argv[1].split('.')[0] + '_LocalEnumerationImport.xlsx')
worksheet = workbook.add_worksheet('Enumerations')
worksheet1 = workbook.add_worksheet('Items')


bold = workbook.add_format({'bold': True})
cell_format = workbook.add_format()

# Titles
worksheet.write(0,0,"Name", bold)
worksheet.write(0,1,"Description", bold)
worksheet1.write(0,0,"Local Enumeration Name", bold)
worksheet1.write(0,1,"Item Name", bold)
worksheet1.write(0,2,"Item Value", bold)

# Start from the first cell. Rows and columns are zero indexed.
row = 1
col = 0

print("Converting...")
# find all enum set names
for enum in root.findall('./EnumerationSet'):

    for enumsetname in enum.findall('./UniqueName'):
        worksheet.write(row,col, enumsetname.text)
        row +=1

# Start from the first cell. Rows and columns are zero indexed.
row = 1
col = 0

for enum in root.findall('./EnumerationSet'):

    for enumsetname in enum.findall('./UniqueName'):
            for member in enum.findall('./Member'):
                    worksheet1.write(row,col, enumsetname.text)
                    for member_name in member.findall('./Name'):
                        worksheet1.write(row,col+1, member_name.text)
                    for member_ordinal in member.findall('./Ordinal'):
                        worksheet1.write(row,col+2, member_ordinal.text)
                    row +=1


workbook.close()
print("Done!\n")