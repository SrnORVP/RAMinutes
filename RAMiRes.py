import pandas as pd, numpy as np, matplotlib.pyplot as plt, main_pandas, os
from sys import platform
from datetime import datetime

strProjID = main_pandas.strProjID
arrProjPath = main_pandas.arrProjPath
strRAMiRes_Input = main_pandas.strRAMiRes_Input
arrParaPath = main_pandas.arrParaPath
arrRAMiCEs_Param = main_pandas.arrRAMiCEs_Param

strRAMiRes_Output = main_pandas.strRAMiRes_Output



#-----------------------------------------------------------------------------------------------------------------------
intRAMheader = 'Element Description'
strEndUnit = 'Units'                                               # Unit of Production (string)
intProdTar = 95                                                     # Production Target (integer)
intPieChartAngle = 60                                               # Start Angle for Pie Chart
#-----------------------------------------------------------------------------------------------------------------------

# Subsystem Criticality
dictDF = pd.read_excel(os.path.join(*arrProjPath, strRAMiRes_Input), sheet_name=None, header=0, dtype='object')
dictDF['Event'].index = dictDF['Event'].set_index(intRAMheader)
dictDF['Element'].index = dictDF['Element'].set_index(intRAMheader)


# Combine the Event and Element DFs
dfEqp = pd.DataFrame({'Tag': dfElem['Info'],
                      'Grp': dfEvet['Group Description'],
                      'Cate': dfEvet['Element Category'],
                      'RelLoss': dfEvet['Relative Loss %'],
                      'AbsLoss': dfEvet['Abs. Loss Mean%']}, index=dfEvet['Element Description'])


# Convert L2 to multi level index
dfGrp = pd.read_excel(xlsName, 'Group', index_col=0)
if '.' not in dfGrp['idxL2']:
    dfGrp = dfGrp.assign(idxL2=dfGrp['idxL1'].apply(str)+'.'+dfGrp['idxL2'].apply(str))

# Get the unique entries of L1 and L2 indexes
dfIdxL1 = dfGrp.loc[:, ['idxL1', 'L1']].set_index('L1').drop_duplicates()
dfIdxL2 = dfGrp.loc[:, ['idxL2', 'L2']].set_index('L2').drop_duplicates()

# Merge the L1 and L2 Groups to main DF
dfEqp = dfEqp.join(dfGrp, on='Grp', how='left')

# Sum based on the Group level and rename the column
dfL1 = dfEqp.groupby('L1').sum().loc[:, ['RelLoss', 'AbsLoss']]
dfL2 = dfEqp.groupby('L2').sum().loc[:, ['RelLoss', 'AbsLoss']]
dfL1 = dfL1.assign(idxL1=dfIdxL1.idxL1).rename({'idxL1': 'idx'}, axis=1)
dfL2 = dfL2.assign(idxL2=dfIdxL2.idxL2).rename(columns={'idxL2': 'idx'})

# Append dfL1 and dfL2 together and sort
dfTot = dfL1.append(dfL2, sort=False)
dfTot['idx'] = dfTot['idx'].astype(float)
dfTot = dfTot.sort_values('idx')
dfTot = dfTot.reset_index().set_index('idx')
dfTot = dfTot.rename({'index': 'Subsystem'}, axis=1)


def float2str(ftInput, dp):
    if ftInput < 1/(10**dp):
        strOutput = f'<{1/(10**dp)}'
    else:
        strOutput = f'{ftInput :.{dp}f}'
    return strOutput


# Convert to str for output, needed for presentation of '<0.01'
dfTot['RelLoss'] = dfTot['RelLoss'].apply(float2str, dp=2)
dfTot['AbsLoss'] = dfTot['AbsLoss'].apply(float2str, dp=2)

# Subsystem Criticality Chart
dfSubCht = dfL1['RelLoss']
arrSubLabel = dfSubCht.index + ' (' + dfSubCht.round(1).apply(str)+'%)'
arrSubLabel = arrSubLabel.values
pltSubsys, axPie = plt.subplots(figsize=(7, 7))
wedgeSubsysPie, lbl = axPie.pie(dfSubCht.values, explode=[0.05]*len(dfSubCht.values), startangle=intPieChartAngle)
bbox_props = dict(boxstyle="round,pad=0.2", ec="k", lw=0.72, alpha=0)
kw = dict(arrowprops=dict(arrowstyle="-"), bbox=bbox_props, zorder=0, va="center")

for i, p in enumerate(wedgeSubsysPie):
    ang = (p.theta2 - p.theta1)/2. + p.theta1
    y = np.sin(np.deg2rad(ang))
    x = np.cos(np.deg2rad(ang))
    horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
    connectionstyle = "angle,angleA=0,angleB={}".format(ang)
    kw["arrowprops"].update({"connectionstyle": connectionstyle})
    axPie.annotate(arrSubLabel[i], xy=(x, y), xytext=(1.2*np.sign(x), 1.2*y),
                   horizontalalignment=horizontalalignment, **kw)

# Individual Equipment Criticality
# Get a DataFrame on Equipment tags
dfTag = dfEqp.loc[:, ['Tag']].reset_index()
dfTag = dfTag.drop_duplicates().set_index('Element Description')

# Sum availability for equipment with same name, add equipment tag to the DataFrame, reorganise
dfInd = dfEqp.groupby('Element Description').sum().loc[:, ['RelLoss', 'AbsLoss']]
dfInd = dfInd.assign(Tag=dfTag['Tag'])
dfInd = dfInd.reindex(columns=['Tag', 'RelLoss', 'AbsLoss'])

# Sort the DataFrame for individual equipment
dfInd = dfInd.sort_values('RelLoss', ascending=False)
dfInd = dfInd[dfInd['RelLoss'] >= 0.1]
for seqCol in [1, 2]:
    dfInd.iloc[:, seqCol] = dfInd.iloc[:, seqCol].apply(float2str, dp=2)

# Production Profile
# Import Production profile, rename and get total potential production
dfProdRaw = pd.read_excel(xlsName, 'Profile')
dfProd = dfProdRaw.rename(columns={f'Production Volume ({strEndUnit})': 'ProdVol',
                                   'Production Efficiency %': 'ProdEff',
                                   f'Losses ({strEndUnit})': 'ProdLoss'})
dfProd = dfProd.assign(TotProd=dfProd['ProdVol'] + dfProd['ProdLoss'])
dfProd['ProTarget'] = intProdTar

# Remove uncessary columns, get daily production
dfProd = dfProd.iloc[:, [0, 1, 2, 7, 3, 8]]
dfProd = dfProd.set_index('Time')
dfProd.iloc[:, 0:3] = np.round(dfProd.iloc[:, 0:3]/365.25, 1)
dfProd.iloc[:, 3] = np.round(dfProd.iloc[:, 3], 1)

dfProdTable = pd.DataFrame({f'Production Volume ({strEndUnit})': dfProd['ProdVol'],
                            f'Losses ({strEndUnit})': dfProd['ProdLoss'],
                            'Production Availability %': dfProd['ProdEff']})
for seqCol in range(len(dfProdTable.columns)):
    dfProdTable.iloc[:, seqCol] = dfProdTable.iloc[:, seqCol].apply(float2str, dp=1)

# Production Profile Chart
# Setup plot area and axis
intFigLen = round(len(dfProd.index)/2, 0)+2
pltProd = plt.figure(figsize=[intFigLen, 4])
axVolBar = pltProd.add_subplot(111)
ftVolYLim = dfProd['TotProd'].max()*1.5
ftVolYLim = round(ftVolYLim, -f"{ftVolYLim}".find(".")+2)
axVolBar.set_xlim(dfProd.index.min()-0.5, dfProd.index.max()+0.5)
axVolBar.set_ylim(0, ftVolYLim)

# plot the line for Total Production and running volume
BarTot = axVolBar.bar(dfProd.index, dfProd['TotProd'], color='yellowgreen',  label='Losses', width=0.8)
BarVol = axVolBar.bar(dfProd.index, dfProd['ProdVol'], color='darkgreen', label='Production',  width=0.8)

# set up Production Volume Bar Chart
axVolBar.set_ylabel(f'Volume ({strEndUnit})')
axVolBar.set_xlabel('Year')
axVolBar.set_xticks(dfProd.index)
axVolBar.spines['top'].set_visible(False)

# plot the line for Production availability and target
axEffLine = axVolBar.twinx()
lnEff = axEffLine.plot(dfProd.index, dfProd['ProdEff'], 'midnightblue')
lnTar = axEffLine.plot(dfProd.index, dfProd['ProTarget'], 'red', label=f'{intProdTar}% Availibility')

# set up Prod Avail Line Chart
axEffLine.set_ylim(50, 115)
axEffLine.set_ylabel('Availabilty (%)')
axEffLine.spines['top'].set_visible(False)

# Setup legend
pltProd.legend(loc=2, bbox_to_anchor=(0.69, 0.9), frameon=False)

# Equipment Type Criticality
dfType = dfEvet.reindex(columns=['Relative Loss %', 'Abs. Loss Mean%', 'Element Category'])
dfType = dfType.rename(columns={'Relative Loss %': 'RelLoss', 'Abs. Loss Mean%': 'AbsLoss', 'Element Category': 'Type'})
dfType = dfType.groupby('Type').sum()
dfType.sort_values(by='RelLoss', ascending=False,  inplace=True)
dfTypeTable = dfType.copy()
for seqCol in range(len(dfTypeTable.columns)):
    dfTypeTable.iloc[:, seqCol] = dfTypeTable.iloc[:, seqCol].apply(float2str,  dp=2)

dfTypeLarge = dfType[dfType['RelLoss'] > 0.5]
dfTypeSmall = dfType[dfType['RelLoss'] <= 0.5]
dfTypeSmall = dfTypeSmall.sum()
dfTypeSmall.name = 'Others'
dfType = dfTypeLarge.append(dfTypeSmall)

# Equipment Type Criticality Chart
intFigLen = round(len(dfType.index), 0)+2
pltType = plt.figure(figsize=[intFigLen, 4])
axTypeBar = pltType.add_axes([0, 0, 1, 1])

ftTypeYLim = dfType['RelLoss'].max()*1.1
ftTypeYLim = round(ftTypeYLim, -f"{ftTypeYLim}".find(".")+2)
axTypeBar.set_ylim(0, ftTypeYLim)

BarType = axTypeBar.bar(dfType.index, dfType['RelLoss'], color='darkgreen', width=0.8)
axTypeBar.set_ylabel(f'Relative Availability %')
axTypeBar.set_xlabel('Equipment Type')
axTypeBar.spines['top'].set_visible(False)

# Export to excel
strTimeNow = datetime.now().strftime("%y%m%d%H%M")
xlsOutput = strPath = f'.{strHash}' + xlsOutput + '-' + strTimeNow
with pd.ExcelWriter(f'{xlsOutput}.xlsx') as xwr:
    dfProdTable.to_excel(xwr, 'ProdProfile')
    dfTot.to_excel(xwr, 'Subsys')
    dfInd.to_excel(xwr, 'Equipment')
    dfTypeTable.to_excel(xwr, 'Type')

pltProd.savefig(f'{xlsOutput}-ProdProfile.png',   bbox_inches='tight')
pltSubsys.savefig(f'{xlsOutput}-Subsystem.png',  bbox_inches='tight',  pad_inches=0.1)
pltType.savefig(f'{xlsOutput}-EquType.png',  bbox_inches='tight')
