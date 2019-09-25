import xml.etree.ElementTree as ET
# tree = ET.parse('control_strategies.axml')
tree = ET.parse('baddi_orig.axml')

root = tree.getroot()
phases=[] # phase name list

# find all phase names
for recipephase in root.findall('./RecipePhase/UniqueName'):
    phases.append(recipephase.text)

n_elem = len(phases) # no of phases

# lists init
param=[[] for _ in range(n_elem)]
strategies=[[] for _ in range(n_elem)]
matrix=[]

# test
print(phases)
print(n_elem)


i = 0 # counter var to cicle through phase index

# XML parsing for parameters, control strategies and associations
for recipephase in root.findall('./RecipePhase'):

    for defparam in recipephase.findall('./DefaultRecipeParameter'):

        for paramname in defparam.findall('./Name'):
            param[i].append(paramname.text)

    for strategy in recipephase.findall('./ControlStrategyAssociations'):
        row = []
        for strategy_name in strategy.findall('./ControlStrategyValue'):
            strategies[i].append(strategy_name.text)


            for j in range(0,len(param[i])):
                check = False
                for strategy_param in strategy.findall('./RecipeParameter/Name'):
                    if param[i][j] in strategy_param.text:

                        check = True

                if check:
                    row.append(1)
                else:
                    row.append(0)

            matrix.append(row)

    i = i + 1 #index increment

i = 0

print(matrix)

for k in range(0,n_elem):
    # csharp code gen
    file = open("./output/" + phases[k],'w')
    file.write("using System;\n")
    file.write("namespace NeoPLM.Neo.DynamicCode {\n")
    file.write("public class PhaseDefinition\n")
    file.write("{\n")
    file.write("	public void RefreshParameters( IPhaseDynamicCodeFacade dynamicCodeFacade )\n")
    file.write("	{\n")
    file.write("		#region *** Begin Dynamic Code ***")
    file.write("\t\n")
    file.write("		string CONTROL_STRATEGY1 = dynamicCodeFacade.GetParameterValueAsString(\"CONTROL_STRATEGY\");")
    file.write("\t\n")
    file.write("        //	=======================================================================================\n")
    file.write("        //	ACTIVATION OF PARAMETERS BASED ON CONTROL STRATEGY ENUMERATION VALUE\n")
    file.write("        //  TO BE USED ONLY FOR PARAMETERS THAT ARE NOT COPIED FROM GENERIC PHASE\n")
    file.write("        //	=======================================================================================\n")
    file.write("\t\n")


    for cs in strategies[k]:

        file.write("        if (CONTROL_STRATEGY1 == \"" + cs + "\")\n")
        file.write("        {\n")

        j = 0
        for par in param[k]:
            file.write("        dynamicCodeFacade.SetParameterEnabledState(\"" + par + "\"," + str(bool(matrix[i][j])).lower() + ");\n")
            j = j + 1
        file.write("        }\n")
        file.write("\t\n")
        i = i + 1

    file.write("		#endregion\n")
    file.write("	}\n")
    file.write("}}")
    file.close()

