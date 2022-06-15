import sys, os

#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------

# General Project Name
# TODO B4
# strProjID = 'B4' + 'RBD'
# strProjID = 'B4' + 'Plant' + '5yr'
# strProjID = 'B4' + 'Plant' + '5yr' + 'No6'
# strProjID = 'B4' + 'Plant' + '5yr' + 'Avg'
# strProjID = 'B4' + 'Prod' + '5yr'
# strProjID = 'B4' + 'Prod' + '5yr' + 'No6'
# strProjID = 'B4' + 'Prod' + '5yr' + 'Avg'
# strProjID = 'B4' + 'Prod' + '20yr'
# strProjID = 'B4' + 'Prod' + '25yr'
# arrProjPath = ["/Users/srnorvpsmac/OneDrive - ERM/0575193 Majnoon/4-Working/2-C&E/4 - CnE - 30Sep21"]
# strRAMiCEs_Input = strProjID + 'RAMinput_05Oct21.xlsx'
# ----------------------------------------------------------------------------------------------------------------------
# TODO Majoon
arrProjPath = ["/Users/srnorvpsmac/OneDrive - ERM/0575193 Majnoon/4-Working/2-C&E/4 - CnE - 30Sep21/Production Availability (for report only)"]

strProjID = 'Majoon' + '_Prod'
strRAMiCEs_Input = 'RAMinput_Prod_for report.xlsx'

# strProjID = 'Majoon' + '_Plant'
# strRAMiCEs_Input = 'RAMinput-Plant-06Oct21.xlsx'
# ----------------------------------------------------------------------------------------------------------------------

# strRAMiRes_Input = strProjID + '-RAMresult' + '.xlsx'

# ----------------------------------------------------------------------------------------------------------------------

isBypassValid = False
isRender = False
isRAMiCEs = True
# isRAMiCEs = False

# ----------------------------------------------------------------------------------------------------------------------

# Relative path and File name of user input file

# Name of Output Identifier
strRAMaros_Output = strProjID + '-RAMaros'
strRAMiCEs_Output = strProjID + '-RAMiCEs'
strRAMiRes_Output = strProjID + '-RAMiRes'

# Relative path and File name of General Param Inputs
arrParaPath = ['.']
strRAMaros_Tmplt = 'RAMaros-Tmplt-' + '03Mar21' + '.xlsm'
arrRAMiCEs_Param = 'RAMiCEs-Param-' + '03Mar21' + '.xlsx'

#-----------------------------------------------------------------------------------------------------------------------
# RAMiRes Param

# Unit of Production as per model (string)
strProdUnit = 'Units'
# Unit of Production as per report (string)
strProdUnitChart = 'MMscfd'
# Production Target (integer)
intProdTar = 95
# the cutoff point for relative availability for Equipment Type Table (number)
numTypeTableCutOff = 0.05

# if subsystem chart label overlap change this param (number)
numSubsysChartParam = -60

strSubsysChartLevel = 'L2'


#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------

if __name__ == '__main__':
    sys.path.append(os.getcwd())
    arrPathShort = os.getcwd().split(os.path.sep)[-2:]
    # intUI = int(input('Which script you would like to run? [RAMiCEs/RAMaros=1]: '))
    intUI=1
    if intUI == 1:
        strPath_Script = os.path.join(*arrPathShort, 'RAMiCEs.py')
        print(f'{strPath_Script} is being ran.\n')

        import RAMiCEs
        RAMiCEs.__name__

    elif intUI == 2:
        pass
    elif intUI == 3:
        pass
    elif intUI == 4:
        pass
    else:
        print('Invalid Input: Script Exit.')
        input('Press any key to exit.')
        exit()

    print(f'{strPath_Script} has ran successfully.')
    # input('Press any key to exit.')

