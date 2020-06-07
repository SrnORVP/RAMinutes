import sys, os

#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------

# General Project Name
strProjID = 'Test'

# Relative path and File name of user input file
arrProjPath = ['..', '01-Test']
strRAMiCEs_Input = strProjID + '-RAMinput-26Jan20' + '.xlsx'
strRAMiRes_Input = strProjID + '-RAMresult-26Jan20' + '.xlsx'

# Name of Output Identifier
strRAMaros_Output = strProjID + '-RAMaros'
strRAMiCEs_Output = strProjID + '-RAMiCEs'
strRAMiRes_Output = strProjID + '-RAMiRes'

# Relative path and File name of General Param Inputs
arrParaPath = ['.']
strRAMaros_Tmplt = 'RAMaros-Tmplt-10Feb20' + '.xlsm'
arrRAMiCEs_Param = 'RAMiCEs-Param-30May20' + '.xlsx'

isBypassValid = False

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
    intUI = int(input('Which script you would like to run? [RAMiCEs/RAMaros=1]: '))
    if intUI == 1:
        strPath_Script = os.path.join(*arrPathShort, 'RAMiCEs.py')
        print(f'{strPath_Script} is being ran.')
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
    input('Press any key to exit.')

