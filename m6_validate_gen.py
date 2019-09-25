import xml.etree.ElementTree as ET
tree = ET.parse('jnj_m6_new.axml')

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

    for defparam in recipephase.findall('./RecipeParameter'):

        for paramname in defparam.findall('./Name'):
            param[i].append(paramname.text)

    check = False

    row = []
    for strategy_param in recipephase.findall('./RecipeParameter'):
        if (strategy_param.find('./RealMin') is not None):
            check = True
            print(str(i) + " " + str(check) + " " + " " + str(strategy_param.find('./Name').text) + " " + str(strategy_param.find('./RealMin').text))
        if (strategy_param.find('./IntegerMin') is not None):
            check = True
            print(str(i) + " " + str(check) + " " + " " + str(strategy_param.find('./Name').text) + " " + str(strategy_param.find('./IntegerMin').text))
        if check:
            row.append(str(strategy_param.find('./Name').text))
            check = False
    matrix.append(row)

    i = i + 1  # index increment

i = 0

print(matrix[0])
print(param)

for k in range(0,n_elem):
    # csharp code gen
    file = open("./validate_m6/validate_" + phases[k],'w')
    file.write("using System;\n")
    file.write("namespace NeoPLM.Neo.DynamicCode {\n")
    file.write("public class PhaseDefinition\n")
    file.write("{\n")
    file.write("	public void Validate( IPhaseDynamicCodeFacade dynamicCodeFacade, ICollection<DynamicCodeValidationResult> validationResults )\n")
    file.write("	{\n")
    file.write("		#region *** Begin Dynamic Code ***")
    file.write("\t\n")

    j = 0
    for par in matrix[i]:

        file.write("			if(!dynamicCodeFacade.GetParameterValueAsDoubleIsInsideBoundsInclusive(\"" + matrix[i][j] + "\"))\n")
        file.write("			{\n")
        file.write("            string errorMessage = \"" + matrix[i][j] + ": The " + matrix[i][j] + " Parameter must be between the Upper and Lower Bound defined in the Phase Definition.\";\n")
        file.write("            validationResults.Add(dynamicCodeFacade.CreateErrorValidationResult(errorMessage, \"" + matrix[i][j] + "\"));\n")
        file.write("            }\n")
        file.write("\n")
        j = j + 1

    file.write("\t\n")
    i = i + 1

    file.write("		#endregion\n")
    file.write("	}\n")
    file.write("}}")
    file.close()

