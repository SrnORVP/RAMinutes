
import pandas as pd, numpy as np, openpyxl as opxl, os, SACusFun as SACF, main_pandas
from anytree import AnyNode, RenderTree, Walker
from datetime import datetime
from itertools import zip_longest


strProjID = main_pandas.strProjID
arrProjPath = main_pandas.arrProjPath
strRAMiCEs_Input = main_pandas.strRAMiCEs_Input
arrParaPath = main_pandas.arrParaPath
strRAMaros_Tmplt = main_pandas.strRAMaros_Tmplt
arrRAMiCEs_Param = main_pandas.arrRAMiCEs_Param
strRAMaros_Output = main_pandas.strRAMaros_Output
strRAMiCEs_output = main_pandas.strRAMiCEs_output
isBypassValid = main_pandas.isBypassValid


arrRAMtype = ['eItemNode','eItemGroup','eItemPBlock','eItemPUnit','eItemUnschElmt','eItemUnschFMode']
dictMarosItem = dict(eItemGroup=4, eItemPBlock=5, eItemPUnit=6, eItemUnschElmt=7, eItemSchElmt=9, eItemUnschFMode=10, eItemSchAct=11, eItemNode=47)

strInputFile = os.path.join(*arrProjPath,strRAMiCEs_Input)
dictExcel_dfRaw = pd.read_excel(strInputFile,None)
dictExcel_dfItem = dictExcel_dfRaw.copy()
dfEqu = dictExcel_dfItem['eItemUnschElmt']
dictExcel_dfItem['eItemUnschElmt'] = dfEqu[dfEqu['equProdLoss'].isin(['yes','Yes','Y','y'])]


# [InputTab, ParamTab, UnityHeader, RAMaroHeader, RAMiCEsHeader, RAMiCEsConfig, RAMiCEsProdLs]
strInputFile = os.path.join(*arrParaPath,arrRAMiCEs_Param)
dictExcel_dfParam = pd.read_excel(arrRAMiCEs_Param,None)


# Check Validity of Uniqueness
arrExcel_srsUnique = []
for strItemTable, dfItemTable in dictExcel_dfItem.items():
    if strItemTable != 'eItemUnschFMode':
        arrExcel_srsUnique.append(dfItemTable.iloc[:,1])
srsUnique = pd.concat(arrExcel_srsUnique,axis=0).value_counts()
if not isBypassValid:
    assert len(srsUnique[srsUnique>1])==0, 'Some items are not unique'


# #### Check Validity of Super and Child
arrExcel_srsSuper = []
arrExcel_srsChild = []
for strItemTable, dfItemTable in dictExcel_dfItem.items():
    if not strItemTable in ['eItemUnschFMode','eItemUnschElmt']:
        arrExcel_srsSuper.append(dfItemTable.iloc[:,0])
        arrExcel_srsChild.append(dfItemTable.iloc[:,1])
    if strItemTable in 'eItemUnschElmt':
        arrExcel_srsSuper.append(dfItemTable.iloc[:,0])
        arrExcel_srsChild.append(dfItemTable.iloc[:,4])
    if strItemTable in 'eItemUnschFMode':
        arrExcel_srsSuper.append(dfItemTable.iloc[:,0])
srsSuper = pd.concat(arrExcel_srsSuper,axis=0).str.split(' #',expand=True)[0].dropna().drop_duplicates()
arrChild = pd.concat(arrExcel_srsChild,axis=0).dropna().drop_duplicates().to_list() + ['Root','root']
srsSuper = srsSuper[~srsSuper.isin(arrChild)]
if not isBypassValid:
    assert len(srsSuper)==0, 'Some Super cannot be found'


# #### Get dfPBU by repeating on pbTrains
dfPBk = dictExcel_dfItem['eItemPBlock']
arrPBU =[]
for strCol in dfPBk.columns:
    arrPBU.append(dfPBk[strCol].repeat(dfPBk['pbTrains']))
dfPBU = pd.concat(arrPBU, axis=1)
dfPBU['pbSuper']=dfPBU['pbUnique']

srsRollingCnt = pd.Series(dfPBU.index).expanding().apply(lambda x: x.value_counts()[x.iloc[-1]])
dfPBU['pbUnique'] = dfPBU['pbUnique'].str.cat(srsRollingCnt.astype(int).astype(str).to_list(),sep=' #')
dfPBU_RAMiCEs = dfPBU.copy()


# #### Get Active Passive Mode from pbuProd and pbuPass

dfPBU['pbuProdSum'] = dfPBU.groupby(dfPBU.index)['pbProd'].transform(pd.Series.cumsum)
dfPBU['pbuMode'] = ~((dfPBU['pbuProdSum'] > 100) & (dfPBU['pbPassive'] == 'Yes'))

dfPBU = dfPBU.rename(columns={'pbSuper':'pbuSuper','pbUnique':'pbuUnique','pbProd':'pbuProd'})
dfPBU = dfPBU.reindex(['pbuSuper','pbuUnique','pbuMode','pbuProd'],axis=1)
dfPBk = dfPBk.reindex(['pbSuper','pbUnique'],axis=1)

dictExcel_dfItem['eItemPBlock'] = dfPBk
dictExcel_dfItem['eItemPUnit'] = dfPBU


# #### Equ Table and FM Table
dfEqu = dictExcel_dfItem['eItemUnschElmt'].reindex(columns=['equSuper','equUnique','equDesc','equType'])
dfFMd = dictExcel_dfItem['eItemUnschElmt'].rename(columns={'equUnique':'fmSuper'})

dfFMd = dfFMd.merge(dictExcel_dfItem['eItemUnschFMode'],'left','equType')
dfFMd['fmRParam2'] = dfFMd.loc[:,['fmRamp','fmLogis','fmMTTR']].astype(np.float).sum(axis=1)
dfFMd['fmRParam1'] = np.round(dfFMd['fmRParam2'] * 0.942,2)
dfFMd['fmUnique'] = pd.Series(range(0,dfFMd.shape[0])).apply((lambda x:'FM' + str(x)))

dictExcel_dfItem['eItemUnschElmt'] = dfEqu
dictExcel_dfItem['eItemUnschFMode'] = dfFMd
dfFMd_RAMiCEs = dfFMd.copy()


# Align Headers and Concate DataFrames
dictRAMname = dictExcel_dfParam['UnityHeader'].pivot(index='strInputExcelTab',columns='strHeaderReplace',values='strHeaderLookup').to_dict('index')
for eItemKey, eItemValue in dictRAMname.items():
    dictRAMname[eItemKey] = {v: k for k, v in eItemValue.items()}
arrRAMarosOrder = dictExcel_dfParam['RAMarosHeader']['strHeaderReplace']

for strDFtype, dfTemp in dictExcel_dfItem.items():
    dfTemp = dfTemp.assign(itemType=dictMarosItem[strDFtype])
    dfTemp = dfTemp.rename(columns=dictRAMname[strDFtype])
    dictExcel_dfItem[strDFtype] = dfTemp.reindex(columns=arrRAMarosOrder)

dfRoot = pd.DataFrame({'itemSuper':['Root'],'itemUnique':['Root'],'itemType':[1]})
dfRAM = pd.concat([dfRoot, *dictExcel_dfItem.values()],axis=0,sort=False,ignore_index=True)


# Render AnyTree
dfTree = dfRAM[~dfRAM['itemType'].isin([10,11])].loc[:,['itemSuper','itemUnique','itemType']].set_index('itemUnique')
dfTree['itemType'] = dfTree['itemType'].replace({47:'Node',4:'Group',5:'ParaB',6:'ParaU',7:'Equip',9:'PMain',1:'Root'})


def make_Tree_From_Unsort(dfInput, dictInput, strItem):
    if not dictInput.get(strItem,None):
        try:
            _ = dictTree[dfInput.at[strItem,'itemSuper']]
        except KeyError:
            make_Tree_From_Unsort(dfInput, dictInput, dfInput.at[strItem,'itemSuper'])
        finally:
            dictInput[strItem] = AnyNode(id=strItem, parent=dictTree[dfInput.at[strItem,'itemSuper']], type=dfInput.at[strItem,'itemType'])


dictTree = {'Root':AnyNode(id='Root', parent=None, type='Root')}
for strItem in dfTree.index:
    make_Tree_From_Unsort(dfTree, dictTree, strItem)
dfRAMaros = dfRAM.set_index('itemUnique').reindex(index=dictTree.keys()).reset_index().reindex(columns=arrRAMarosOrder)
dfRAMaros = pd.concat([dfRAMaros,dfRAM[dfRAM['itemType'].isin([10,11])]],axis=0)


# Get Tree Heirarchy to Dict
srsEquName = dfRAM[dfRAM['itemType'].isin([7])]['itemUnique']
objWalker = Walker()
dictTWalks_ID = dict()
dictTWalks_Tp = dict()
for elemEquName in srsEquName:
    tupTWalk = objWalker.walk(dictTree['Root'],dictTree[elemEquName])[2]
    dictTWalks_ID[elemEquName] = [nodeElem.id for nodeElem in tupTWalk[:-1]]
    dictTWalks_Tp[elemEquName] = [nodeElem.type for nodeElem in tupTWalk[:-1]]


# Org Tree Dict to DataFrame
arrTWalks_ID = [arrXsect for arrXsect in zip_longest(*dictTWalks_ID.values())]
arrTWalks_Tp = [arrXsect for arrXsect in zip_longest(*dictTWalks_Tp.values())]
dfTWalks_strEquName = pd.DataFrame(arrTWalks_ID,columns=dictTWalks_ID.keys()).T
dfTWalks_strEquType = pd.DataFrame(arrTWalks_Tp,columns=dictTWalks_Tp.keys()).T
dfTWalks = pd.concat([dfTWalks_strEquName,dfTWalks_strEquType],axis=1)
dfTWalks.columns = dfTWalks.columns.astype(str).str.cat(['Level']*len(dfTWalks.columns),sep='-')
dfTWalks_RAMiCEs = dfTWalks.iloc[:,[0,1]]


# Get Configuration, Equipment Type and Failure Rate for each equipment
dfRAMiCEs = dfFMd_RAMiCEs.merge(dfPBU_RAMiCEs,how='left',left_on='equSuper',right_on='pbUnique')
dfRAMiCEs = dfRAMiCEs.merge(dictExcel_dfParam['RAMiCEsConfig'],how='left',on=['pbTrains','pbProd'])
dfRAMiCEs = dfRAMiCEs.merge(dictExcel_dfParam['RAMiCEsProdLs'],how='left',on=['pbTrains','pbProd'])
dfRAMiCEs = dfRAMiCEs.merge(dfTWalks_RAMiCEs,how='left',left_on='fmSuper',right_index=True)
dfRAMiCEs['equConfig'] = dfRAMiCEs['equConfig'].fillna('1x100%') 
dfRAMiCEs['equProdLs'] = dfRAMiCEs['equProdLs'].fillna('0%')


# CleanUp RAMiCEs
dfRAMiCEs = dfRAMiCEs.reindex(dictExcel_dfParam['RAMiCEsHeader']['strHeaderLookup'],axis=1)
dictHeadMap_RAMiCEs = dictExcel_dfParam['RAMiCEsHeader'].set_index('strHeaderLookup').to_dict()['strHeaderReplace']
dfRAMiCEs = dfRAMiCEs.rename(columns=dictHeadMap_RAMiCEs)
dfRAMiCEs.columns.name=''


# Open RAMaros and Export
xlwbMaro = opxl.load_workbook(strRAMaros_Tmplt, read_only=False, keep_vba=True)
xlwsMaro = xlwbMaro['RAMaros']
SACF.write_pdDF_to_opxlWS(xlwsMaro, dfRAMaros, numColOffSet=1, numRowOffSet=1, isIndexWrite=False)

strTimeNow = datetime.now().strftime("%d%b%y")
strRAMaros_Output = os.path.join(*arrProjPath, strRAMaros_Output)
xlwbMaro.save(f'{strRAMaros_Output}.xlsm')


# Export to Excel
strRAMiCEs_output = os.path.join(*arrProjPath, strRAMiCEs_output)
with pd.ExcelWriter(f'{strRAMiCEs_output}.xlsx') as xwr:
    dfTWalks.to_excel(xwr,'Layers')
    dfRAMiCEs.to_excel(xwr,'CnE')

print()
print(f'{"<>"*15} Check rendered tree generated from RAMiCEs {"<>"*15}')
print()
print(RenderTree(dictTree['Root']))
print()
print(f'{"<>"*52}')
print()