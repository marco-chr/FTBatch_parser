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
workbook = xlsxwriter.Workbook('baddi_parameter_bounds.xlsx')
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

        for recipeparam in recipe.findall('./RecipeParameter'):
            worksheet.write(row,col+1, recipephase.text)
            for recipeparamname in recipeparam.findall('./Name'):
                worksheet.write(row,col, recipephase.text + recipeparamname.text)
                worksheet.write(row,col+3, recipeparamname.text)
            for recipeparam_realmin in recipeparam.findall('./RealMin'):
                worksheet.write(row,col+4, recipeparam_realmin.text)
            for recipeparam_realmax in recipeparam.findall('./RealMax'):
                worksheet.write(row,col+5, recipeparam_realmax.text)
            for recipeparam_realdef in recipeparam.findall('./RealDefault'):
                worksheet.write(row,col+6, recipeparam_realdef.text)
            for recipeparam_intmin in recipeparam.findall('./IntegerMin'):
                worksheet.write(row,col+4, recipeparam_intmin.text)
            for recipeparam_intmax in recipeparam.findall('./IntegerMax'):
                worksheet.write(row,col+5, recipeparam_intmax.text)
            for recipeparam_intdef in recipeparam.findall('./IntegerDefault'):
                worksheet.write(row,col+6, recipeparam_intdef.text)
            for recipeparam_eu in recipeparam.findall('./EngineeringUnits'):
                worksheet.write(row,col+7, recipeparam_eu.text)
            row +=1
        for reportparam in recipe.findall('./ReportParameter'):
            worksheet.write(row,col+1, recipephase.text)
            for recipeparamname in reportparam.findall('./Name'):
                worksheet.write(row,col, recipephase.text + recipeparamname.text)
                worksheet.write(row,col+3, recipeparamname.text)
            for recipeparam_realmin in reportparam.findall('./RealMin'):
                worksheet.write(row,col+4, recipeparam_realmin.text)
            for recipeparam_realmax in reportparam.findall('./RealMax'):
                worksheet.write(row,col+5, recipeparam_realmax.text)
            for recipeparam_realdef in reportparam.findall('./RealDefault'):
                worksheet.write(row,col+6, recipeparam_realdef.text)
            for recipeparam_intmin in reportparam.findall('./IntegerMin'):
                worksheet.write(row,col+4, recipeparam_intmin.text)
            for recipeparam_intmax in reportparam.findall('./IntegerMax'):
                worksheet.write(row,col+5, recipeparam_intmax.text)
            for recipeparam_intdef in reportparam.findall('./IntegerDefault'):
                worksheet.write(row,col+6, recipeparam_intdef.text)
            for recipeparam_eu in reportparam.findall('./EngineeringUnits'):
                worksheet.write(row,col+7, recipeparam_eu.text)
            row +=1


        for strategy in recipe.findall('./ControlStrategyAssociations'):
            for strategy_name in strategy.findall('./ControlStrategyValue'):

                for strategy_param in strategy.findall('./RecipeParameter'):
                    worksheet.write(row,col+1, recipephase.text)
                    worksheet.write(row,col+2, strategy_name.text)
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

                for strategy_param in strategy.findall('./DefaultRecipeParameter'):
                    worksheet.write(row,col+1, recipephase.text)
                    worksheet.write(row,col+2, strategy_name.text)
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