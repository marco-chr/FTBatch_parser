import xlsxwriter
import xml.etree.ElementTree as ET
import sys

# Open Area XML file and look for root
# Delete namespace from root tag before launching this script
print("Opening axml file")
tree = ET.parse(sys.argv[1])
root = tree.getroot()
phases=[] # phase name list

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('m6_parameter_bounds.xlsx')
worksheet = workbook.add_worksheet('Limits')
# Create formats
bold = workbook.add_format({'bold': True})
cell_format = workbook.add_format()


# Titles
worksheet.write(0,0,"KEY", bold)
worksheet.write(0,1,"PHASE", bold)
worksheet.write(0,2,"CONTROL STRATEGY", bold)
worksheet.write(0,3,"PARAMETER", bold)
worksheet.write(0,4,"MIN", bold)
worksheet.write(0,5,"MAX", bold)
worksheet.write(0,6,"DEFAULT", bold)
worksheet.write(0,7,"EU", bold)

# Start from the first cell. Rows and columns are zero indexed.
row = 1
col = 0

# find all phase names
print("Converting...")
for recipe in root.findall('./RecipePhase'):

    for recipephase in recipe.findall('./UniqueName'):

        for strategy_param in recipe.findall('./RecipeParameter'):
            worksheet.write(row,col+1, recipephase.text)
            for strategy_param_name in strategy_param.findall('./Name'):
                worksheet.write(row,col, recipephase.text + strategy_param_name.text)
                worksheet.write(row,col+3, strategy_param_name.text)
            for strategy_param_min in strategy_param.findall('./RealMin'):
                worksheet.write(row,col+4, strategy_param_min.text)
            for strategy_param_max in strategy_param.findall('./RealMax'):
                worksheet.write(row,col+5, strategy_param_max.text)
            for strategy_param_def in strategy_param.findall('./RealDefault'):
                worksheet.write(row,col+6, strategy_param_def.text)
            for strategy_param_min in strategy_param.findall('./IntegerMin'):
                worksheet.write(row,col+4, strategy_param_min.text)
            for strategy_param_max in strategy_param.findall('./IntegerMax'):
                worksheet.write(row,col+5, strategy_param_max.text)
            for strategy_param_def in strategy_param.findall('./IntegerDefault'):
                worksheet.write(row,col+6, strategy_param_def.text)
            for strategy_param_eu in strategy_param.findall('./EngineeringUnits'):
                worksheet.write(row,col+7, strategy_param_eu.text)
            row +=1


workbook.close()
print("Done!")