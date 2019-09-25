# Convert phase definitions contained in BES export .axml file to XLS Import file for NEOPLM
# Release 0.0 - 31 Oct 2018
#
# Usage python phase_import_gen.py area.axml bounds.xlsx
# This script generate only an INTERMEDIATE workbook, do not use the output file for Bulk Import
#
# Before Launching this script use: 'python baddi_bounds_gen.py area.axml' or 'python m3_bounds_gen.py area.axml'

# TODO: Target Engineering units conversion
# TODO: XML file import withou root attrib deletion

import xlsxwriter
import xml.etree.ElementTree as ET
import sys
import openpyxl

# Open Bounds and EU workbook
minmax_wb = openpyxl.load_workbook(sys.argv[2], read_only=True)
minmax_ws = minmax_wb['Limits']

# Open Area XML file and look for root
# Delete namespace from root tag before launching this script
print("Opening axml file")
tree = ET.parse(sys.argv[1])
root = tree.getroot()
namespaces = {}

# This dictionary is used to convert Rockwell Datatypes to NEOPLM Datatypes
typedict = {'Real':'Double', 'Integer':'Integer', 'Boolean':'Boolean', 'String':'String'}

# Create a workbook and add a worksheet.
print("Creating workbook")
workbook = xlsxwriter.Workbook(sys.argv[1].split('.')[0] + '_LocalPhaseDefinition_INTERMEDIATE.xlsx')
worksheet = workbook.add_worksheet('Phase Definitions')
worksheet1 = workbook.add_worksheet('Parameters')
worksheet2 = workbook.add_worksheet('FDS_Scope')
worksheet3 = workbook.add_worksheet('Min_Max')

print("Copying bounds and EU to Min_Max worksheet")
n = m = 0
for row in minmax_ws.rows:
    m = 0
    for cell in row:
        worksheet3.write(n,m,cell.value)
        m += 1
    n += 1

# create formats
bold = workbook.add_format({'bold': True})
cell_format = workbook.add_format()
temp = workbook.add_format({'bold': True, 'font_color': 'red'})

# Titles for first worksheet
worksheet.write(0,0,"Phase Def Name", bold)
worksheet.write(0,1,"Phase SCOPE", temp)
worksheet.write(0,2,"Lifecycle", bold)
worksheet.write(0,3,"Description", bold)
worksheet.write(0,4,"PCS Name", bold)

# Titles for second worksheet
worksheet1.write(0,0,"Phase Def Name", bold)
worksheet1.write(0,1,"LifeCycle", bold)
worksheet1.write(0,2,"Param Name", bold)
worksheet1.write(0,3,"Sequence No", bold)
worksheet1.write(0,4,"Description", bold)
worksheet1.write(0,5,"Question", bold)
worksheet1.write(0,6,"Group Name", bold)
worksheet1.write(0,7,"Classification", bold)
worksheet1.write(0,8,"PCS Name", bold)
worksheet1.write(0,9,"Dimension or Enumeration Type", bold)
worksheet1.write(0,10,"Dimension or Enumeration", bold)
worksheet1.write(0,11,"Target Engineering Unit", bold)
worksheet1.write(0,12,"Local Enum Name As Value", bold)
worksheet1.write(0,13,"Is Multi Value", bold)
worksheet1.write(0,14,"Is Visible", bold)
worksheet1.write(0,15,"Is Editable", bold)
worksheet1.write(0,16,"Is Recipe Parameter", bold)
worksheet1.write(0,17,"Recipe Parameter Type", bold)
worksheet1.write(0,18,"Parameter UI Control", bold)
worksheet1.write(0,19,"Recipe Defer Name", bold)
worksheet1.write(0,20,"Default Recipe Defer Level", bold)
worksheet1.write(0,21,"FDS Scope", temp)

# Start from the first cell. Rows and columns are zero indexed.
row = 1
col = 0


# find all phase names
print("Converting...")
for recipe in root.findall('./RecipePhase'):

    for recipename in recipe.findall('./UniqueName'):
        worksheet.write(row,col, recipename.text)
        worksheet.write(row,col+2, "InProcess")
        worksheet.write(row,col+3, recipename.text + " Phase Class")
        worksheet.write(row,col+4, recipename.text)
        row +=1

row = 1
col = 0

for recipe in root.findall('./RecipePhase'):

    for recipename in recipe.findall('./UniqueName'):
        i = 0
        # phase parameter for phases wo control strategies
        for recipeparam in recipe.findall('./RecipeParameter'):
                    for param_name in recipeparam.findall('./Name'):
                        worksheet1.write(row,col, recipename.text)
                        worksheet1.write(row,col+1, "InProcess")
                        worksheet1.write(row,col+2, param_name.text)
                        worksheet1.write(row,col+3, i)
                        worksheet1.write(row,col+7, "Required")
                        worksheet1.write(row,col+8, param_name.text)
                    # dimension
                    if recipeparam.find('EnumerationSetName') is None:
                        worksheet1.write(row,col+5, "Specify " + param_name.text)
                        worksheet1.write(row,col+9, "Dimension")
                        for dimtype in recipeparam.findall('./Type'):
                            worksheet1.write(row,col+10, typedict[dimtype.text])
                        for engunit in recipeparam.findall('./EngineeringUnits'):
                            worksheet1.write(row,col+11, engunit.text)
                        worksheet1.write(row,col+12, "False")
                        worksheet1.write(row,col+13, "False")
                        worksheet1.write(row,col+14, "True")
                        worksheet1.write(row,col+15, "True")
                        worksheet1.write(row,col+16, "True")
                        worksheet1.write(row,col+17, "None")
                        worksheet1.write(row,col+18, "UOMParameterControl")
                        worksheet1.write(row,col+20, "Phase")
                    # enumeration
                    else:
                        worksheet1.write(row,col+5, "Select " + param_name.text)
                        worksheet1.write(row,col+9, "LocalEnum")
                        for enumname in recipeparam.findall('./EnumerationSetName'):
                            worksheet1.write(row,col+10, enumname.text)
                        worksheet1.write(row,col+12, "True")
                        worksheet1.write(row,col+13, "False")
                        worksheet1.write(row,col+14, "True")
                        worksheet1.write(row,col+15, "True")
                        worksheet1.write(row,col+16, "True")
                        worksheet1.write(row,col+17, "None")
                        worksheet1.write(row,col+18, "EnumeratedParameterControl")
                        worksheet1.write(row,col+20, "Phase")

                    i += 1
                    row +=1
        # report parameter for phases wo control strategies
        for recipeparam in recipe.findall('./ReportParameter'):
                    for param_name in recipeparam.findall('./Name'):
                        worksheet1.write(row,col, recipename.text)
                        worksheet1.write(row,col+1, "InProcess")
                        worksheet1.write(row,col+2, param_name.text)
                        worksheet1.write(row,col+3, i)
                        worksheet1.write(row,col+7, "Required")
                        worksheet1.write(row,col+8, param_name.text)
                    # dimension
                    if recipeparam.find('EnumerationSetName') is None:
                        worksheet1.write(row,col+5, param_name.text + " Report Value")
                        worksheet1.write(row,col+9, "Dimension")
                        for dimtype in recipeparam.findall('./Type'):
                            worksheet1.write(row,col+10, typedict[dimtype.text])
                        for engunit in recipeparam.findall('./EngineeringUnits'):
                            worksheet1.write(row,col+11, engunit.text)
                        worksheet1.write(row,col+12, "False")
                        worksheet1.write(row,col+13, "False")
                        worksheet1.write(row,col+14, "True")
                        worksheet1.write(row,col+15, "False")
                        worksheet1.write(row,col+16, "False")
                        worksheet1.write(row,col+17, "None")
                        worksheet1.write(row,col+18, "UOMParameterControl")
                        worksheet1.write(row,col+20, "Phase")
                    # enumeration
                    else:
                        worksheet1.write(row,col+5, param_name.text + " Report Value")
                        worksheet1.write(row,col+9, "LocalEnum")
                        for enumname in recipeparam.findall('./EnumerationSetName'):
                            worksheet1.write(row,col+10, enumname.text)
                        worksheet1.write(row,col+12, "True")
                        worksheet1.write(row,col+13, "False")
                        worksheet1.write(row,col+14, "True")
                        worksheet1.write(row,col+15, "True")
                        worksheet1.write(row,col+16, "False")
                        worksheet1.write(row,col+17, "None")
                        worksheet1.write(row,col+18, "EnumeratedParameterControl")
                        worksheet1.write(row,col+20, "Phase")

                    i += 1
                    row +=1
        # phase parameter for phases with control strategies
        for recipeparam in recipe.findall('./DefaultRecipeParameter'):
                    for param_name in recipeparam.findall('./Name'):
                        worksheet1.write(row,col, recipename.text)
                        worksheet1.write(row,col+1, "InProcess")
                        worksheet1.write(row,col+2, param_name.text)
                        worksheet1.write(row,col+3, i)
                        worksheet1.write(row,col+7, "Required")
                        worksheet1.write(row,col+8, param_name.text)
                    # dimension
                    if recipeparam.find('EnumerationSetName') is None:
                        worksheet1.write(row,col+5, "Specify " + param_name.text)
                        worksheet1.write(row,col+9, "Dimension")
                        for dimtype in recipeparam.findall('./Type'):
                            worksheet1.write(row,col+10, typedict[dimtype.text])
                        for engunit in recipeparam.findall('./EngineeringUnits'):
                            worksheet1.write(row,col+11, engunit.text)
                        worksheet1.write(row,col+12, "False")
                        worksheet1.write(row,col+13, "False")
                        worksheet1.write(row,col+14, "True")
                        worksheet1.write(row,col+15, "True")
                        worksheet1.write(row,col+16, "True")
                        worksheet1.write(row,col+17, "None")
                        worksheet1.write(row,col+18, "UOMParameterControl")
                        worksheet1.write(row,col+20, "Phase")
                    # enumeration
                    else:
                        worksheet1.write(row,col+5, "Select " + param_name.text)
                        worksheet1.write(row,col+9, "LocalEnum")
                        for enumname in recipeparam.findall('./EnumerationSetName'):
                            worksheet1.write(row,col+10, enumname.text)
                        worksheet1.write(row,col+12, "True")
                        worksheet1.write(row,col+13, "False")
                        worksheet1.write(row,col+14, "True")
                        worksheet1.write(row,col+15, "True")
                        worksheet1.write(row,col+16, "True")
                        worksheet1.write(row,col+17, "None")
                        worksheet1.write(row,col+18, "EnumeratedParameterControl")
                        worksheet1.write(row,col+20, "Phase")
                    i += 1
                    row +=1
        # report parameter for phases with control strategies
        for recipeparam in recipe.findall('./DefaultReportParameter'):
                    for param_name in recipeparam.findall('./Name'):
                        worksheet1.write(row,col, recipename.text)
                        worksheet1.write(row,col+1, "InProcess")
                        worksheet1.write(row,col+2, param_name.text)
                        worksheet1.write(row,col+3, i)
                        worksheet1.write(row,col+7, "Required")
                        worksheet1.write(row,col+8, param_name.text)
                    # dimension
                    if recipeparam.find('EnumerationSetName') is None:
                        worksheet1.write(row,col+5, param_name.text + " Report Value")
                        worksheet1.write(row,col+9, "Dimension")
                        for dimtype in recipeparam.findall('./Type'):
                            worksheet1.write(row,col+10, typedict[dimtype.text])
                        for engunit in recipeparam.findall('./EngineeringUnits'):
                            worksheet1.write(row,col+11, engunit.text)
                        worksheet1.write(row,col+12, "False")
                        worksheet1.write(row,col+13, "False")
                        worksheet1.write(row,col+14, "True")
                        worksheet1.write(row,col+15, "False")
                        worksheet1.write(row,col+16, "False")
                        worksheet1.write(row,col+17, "None")
                        worksheet1.write(row,col+18, "UOMParameterControl")
                        worksheet1.write(row,col+20, "Phase")
                    # enumeration
                    else:
                        worksheet1.write(row,col+5, param_name.text + " Report Value")
                        worksheet1.write(row,col+9, "LocalEnum")
                        for enumname in recipeparam.findall('./EnumerationSetName'):
                            worksheet1.write(row,col+10, enumname.text)
                        worksheet1.write(row,col+12, "True")
                        worksheet1.write(row,col+13, "False")
                        worksheet1.write(row,col+14, "True")
                        worksheet1.write(row,col+15, "True")
                        worksheet1.write(row,col+16, "False")
                        worksheet1.write(row,col+17, "None")
                        worksheet1.write(row,col+18, "EnumeratedParameterControl")
                        worksheet1.write(row,col+20, "Phase")

                    i += 1
                    row +=1


workbook.close()
print("Done!\n")
# end