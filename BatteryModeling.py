# -*- coding: utf-8 -*-
"""
Created on Thu Feb 23 18:10:38 2023

@author: c00613768
"""
import numpy as np
import pandas as pd

def SOC_OCV_Table_Extraction(RatedCapacity, lightDscCurrent, heavyDscCurrent, 
                              testFileLocation, testFileName, testSheetName, useCols, 
                              lightDscStartIdx, lightDscEndIdx, heavyDscStartIdx, heavyDscEndIdx, 
                              outputTableLocation, outputTableName):
    
    testData = pd.read_excel(testFileLocation + '\\' + testFileName, sheet_name=testSheetName, header=0, usecols=useCols)    

    lightDscData = testData.loc[lightDscStartIdx:lightDscEndIdx, ['Current(A)', 'Voltage(V)', 'Discharge_Capacity(Ah)']]
    heavyDscData = testData.loc[heavyDscStartIdx:, ['Current(A)', 'Voltage(V)', 'Discharge_Capacity(Ah)']]

    lightStartDscCap = lightDscData.loc[lightDscStartIdx,'Discharge_Capacity(Ah)']
    lightEndDscCap = lightDscData.loc[lightDscEndIdx,'Discharge_Capacity(Ah)']
    lightDscCap = lightEndDscCap - lightStartDscCap

    heavyStartDscCap = heavyDscData.loc[heavyDscStartIdx,'Discharge_Capacity(Ah)']
    heavyEndDscCap = heavyDscData.loc[heavyDscEndIdx,'Discharge_Capacity(Ah)']
    heavyDscCap = heavyEndDscCap - heavyStartDscCap

    lightDscData.index = range(0, lightDscEndIdx + 1 - lightDscStartIdx, 1)
    heavyDscData.index = range(0, heavyDscEndIdx + 1 - heavyDscStartIdx, 1)

    lightSOC_Array = []
    for index, row in lightDscData.iterrows():
        SOC = 100 - 100. * (row['Discharge_Capacity(Ah)'] - lightStartDscCap) / lightDscCap
        lightSOC_Array.append(SOC)
    heavySOC_Array = []
    for index, row in heavyDscData.iterrows():
        SOC = 100 - 100. * (row['Discharge_Capacity(Ah)'] - heavyStartDscCap) / heavyDscCap
        heavySOC_Array.append(SOC)
        
    lightV_I_SOC_Data = pd.concat([lightDscData, pd.DataFrame(lightSOC_Array, columns=['SOC'])], axis=1)[['Voltage(V)','Current(A)','SOC']]
    heavyV_I_SOC_Data = pd.concat([heavyDscData, pd.DataFrame(heavySOC_Array, columns=['SOC'])], axis=1)[['Voltage(V)','Current(A)','SOC']]

    SOC_V_I_R_Table = pd.DataFrame(columns=['SOC','lightV','heavyV','OCV', 'R'])
    for intSOC in range(101):
        SOC_V_I_R_Table.at[intSOC, 'SOC'] = intSOC    
        lightV = lightV_I_SOC_Data.iloc[(lightV_I_SOC_Data['SOC'] - intSOC).abs().argsort()[lightV_I_SOC_Data.index[0]], 0]
        lightI = lightV_I_SOC_Data.iloc[(lightV_I_SOC_Data['SOC'] - intSOC).abs().argsort()[lightV_I_SOC_Data.index[0]], 1]
        heavyV = heavyV_I_SOC_Data.iloc[(heavyV_I_SOC_Data['SOC'] - intSOC).abs().argsort()[heavyV_I_SOC_Data.index[0]], 0]
        heavyI = heavyV_I_SOC_Data.iloc[(heavyV_I_SOC_Data['SOC'] - intSOC).abs().argsort()[heavyV_I_SOC_Data.index[0]], 1]
        SOC_V_I_R_Table.at[intSOC, 'lightV'] = lightV
        SOC_V_I_R_Table.at[intSOC, 'lightI'] = lightI
        SOC_V_I_R_Table.at[intSOC, 'heavyV'] = heavyV
        SOC_V_I_R_Table.at[intSOC, 'heavyI'] = heavyI
        #To avoid singular matrix error
        if lightI == heavyI:
            if lightI == 0:
                SOC_V_I_R_Table.at[intSOC, 'OCV'] = lightV
        else:            
            SOC_V_I_R_Table.at[intSOC, 'OCV'] = np.dot(np.matrix([[1, lightI],[1, heavyI]]).I, np.array([lightV, heavyV]))[0, 0]
            SOC_V_I_R_Table.at[intSOC, 'R'] = np.dot(np.matrix([[1, lightI],[1, heavyI]]).I, np.array([lightV, heavyV]))[0, 1]

    SOC_OCV_Table = SOC_V_I_R_Table.loc[:,['SOC','OCV']]
    SOC_OCV_Table.to_excel(outputTableLocation + '\\' + outputTableName)
    
    # testData = pd.read_excel(testFileLocation + '\\' + testFileName, sheet_name=testSheetName, header=0, usecols=useCols)    
    
    # lightDscData = testData.loc[lightDscStartIdx:lightDscEndIdx, ['Current(A)', 'Voltage(V)', 'Discharge_Capacity(Ah)']]
    # heavyDscData = testData.loc[heavyDscStartIdx:, ['Current(A)', 'Voltage(V)', 'Discharge_Capacity(Ah)']]
    
    # lightStartDscCap = lightDscData.loc[lightDscStartIdx,'Discharge_Capacity(Ah)']
    # lightEndDscCap = lightDscData.loc[lightDscEndIdx,'Discharge_Capacity(Ah)']
    # lightDscCap = lightEndDscCap - lightStartDscCap

    # heavyStartDscCap = heavyDscData.loc[heavyDscStartIdx,'Discharge_Capacity(Ah)']
    # heavyEndDscCap = heavyDscData.loc[heavyDscEndIdx,'Discharge_Capacity(Ah)']
    # heavyDscCap = heavyEndDscCap - heavyStartDscCap
    
    # lightDscData.index = range(0, lightDscEndIdx + 1 - lightDscStartIdx, 1)
    # heavyDscData.index = range(0, heavyDscEndIdx + 1 - heavyDscStartIdx, 1)
    
    # lightSOC_Array = []
    # for index, row in lightDscData.iterrows():
    #     SOC = 100 - 100. * (row['Discharge_Capacity(Ah)'] - lightStartDscCap) / lightDscCap
    #     lightSOC_Array.append(SOC)
    # heavySOC_Array = []
    # for index, row in heavyDscData.iterrows():
    #     SOC = 100 - 100. * (row['Discharge_Capacity(Ah)'] - heavyStartDscCap) / heavyDscCap
    #     heavySOC_Array.append(SOC)
        
    # lightV_SOC_Data = pd.concat([lightDscData, pd.DataFrame(lightSOC_Array, columns=['SOC'])], axis=1)[['Voltage(V)','SOC']]
    # heavyV_SOC_Data = pd.concat([heavyDscData, pd.DataFrame(heavySOC_Array, columns=['SOC'])], axis=1)[['Voltage(V)','SOC']]
    
    # SOC_V_R_Table = pd.DataFrame(columns=['SOC','lightV','heavyV','OCV', 'R'])
    # for intSOC in range(101):
    #     SOC_V_R_Table.at[intSOC, 'SOC'] = intSOC
    #     SOC_V_R_Table.at[intSOC, 'lightV'] = lightV_SOC_Data.iloc[(lightV_SOC_Data['SOC'] - intSOC).abs().argsort()[lightV_SOC_Data.index[0]], 0]
    #     SOC_V_R_Table.at[intSOC, 'heavyV'] = heavyV_SOC_Data.iloc[(heavyV_SOC_Data['SOC'] - intSOC).abs().argsort()[heavyV_SOC_Data.index[0]], 0]
    #     SOC_V_R_Table.at[intSOC, 'OCV'] = np.dot(np.matrix([[1, -lightDscCurrent],[1, -heavyDscCurrent]]).I, np.array([SOC_V_R_Table.at[intSOC, 'lightV'], SOC_V_R_Table.at[intSOC, 'heavyV']]))[0, 0]
    #     SOC_V_R_Table.at[intSOC, 'R'] = np.dot(np.matrix([[1, -lightDscCurrent],[1, -heavyDscCurrent]]).I, np.array([SOC_V_R_Table.at[intSOC, 'lightV'], SOC_V_R_Table.at[intSOC, 'heavyV']]))[0, 1]
    
    # SOC_OCV_Table = SOC_V_R_Table.loc[:,['SOC','OCV']]
    # SOC_OCV_Table.to_excel(outputTableLocation + '\\' + outputTableName)
    
    return SOC_OCV_Table


RatedCapacity = 4.360
lightDscCurrent = RatedCapacity / 40.
heavyDscCurrent = 2 * lightDscCurrent

lightDscStartIdx_DS = 7250
lightDscEndIdx_DS = 145690
heavyDscStartIdx_DS = 152862
heavyDscEndIdx_DS = 222212
testFileLocation_DS = r'D:\Chris\Projects\Coulomb Meter\具体项目\手机\Cetus\Cetus-\协助RT攻关\BT231_25C_OCV_DS20230220_2023_02_20_113444'
testFileName_DS = r'BT231_25C_OCV_DS20230220_Channel_1_Wb_1.xlsx'
testSheetName_DS = r'Channel-1_1'
useCols_DS = 'A:H,J:K'
outputTableLocation_DS = r'D:\Chris\Projects\Coulomb Meter\具体项目\手机\Cetus\Cetus-\协助RT攻关\BT231_25C_OCV_DS20230220_2023_02_20_113444'
outputTableName_DS = r'SOC_OCV_Table_25C_DS.xlsx'

lightDscStartIdx_XWD = 7875
lightDscEndIdx_XWD = 146514
heavyDscStartIdx_XWD = 153670
heavyDscEndIdx_XWD = 223047
testFileLocation_XWD = r'D:\Chris\Projects\Coulomb Meter\具体项目\手机\Cetus\Cetus-\协助RT攻关\BT231_25C_OCV_XWD20230220_2023_02_20_115258'
testFileName_XWD = r'BT231_25C_OCV_XWD20230220_Channel_3_Wb_1.xlsx'
testSheetName_XWD = r'Channel-3_1'
useCols_XWD = ':H,J:K'
outputTableLocation_XWD = r'D:\Chris\Projects\Coulomb Meter\具体项目\手机\Cetus\Cetus-\协助RT攻关\BT231_25C_OCV_XWD20230220_2023_02_20_115258'
outputTableName_XWD = r'SOC_OCV_Table_25C_XWD.xlsx'

SOC_OCV_DS_Table = SOC_OCV_Table_Extraction(RatedCapacity, lightDscCurrent, heavyDscCurrent, 
                                            testFileLocation_DS, testFileName_DS, testSheetName_DS, useCols_DS, 
                                            lightDscStartIdx_DS, lightDscEndIdx_DS, heavyDscStartIdx_DS, heavyDscEndIdx_DS, 
                                            outputTableLocation_DS, outputTableName_DS)

SOC_OCV_XWD_Table = SOC_OCV_Table_Extraction(RatedCapacity, lightDscCurrent, heavyDscCurrent, 
                                            testFileLocation_XWD, testFileName_XWD, testSheetName_XWD, useCols_XWD, 
                                            lightDscStartIdx_XWD, lightDscEndIdx_XWD, heavyDscStartIdx_XWD, heavyDscEndIdx_XWD, 
                                            outputTableLocation_XWD, outputTableName_XWD)

RatedCapacity = 0.530
lightDscCurrent = RatedCapacity / 40.
heavyDscCurrent = 2 * lightDscCurrent

lightDscStartIdx = 14499
lightDscEndIdx = 42041
heavyDscStartIdx = 43721
heavyDscEndIdx = 57869
testFileLocation = r'D:\Chris\Projects\Coulomb Meter\AFE项目\算法验证\Battery\BT234 & BT238\BT234\BT234\New added\OCV-SOC Calculation'
testFileName = r'BT234-RT-OCV-1_Channel_39.xlsx'
testSheetName = r'Channel_39_1'
useCols = ':G,H:I'
outputTableLocation = testFileLocation
outputTableName = r'SOC_OCV_Table_25C.xlsx'

SOC_OCV_Table = SOC_OCV_Table_Extraction(RatedCapacity, lightDscCurrent, heavyDscCurrent, 
                                         testFileLocation, testFileName, testSheetName, useCols, 
                                         lightDscStartIdx, lightDscEndIdx, heavyDscStartIdx, heavyDscEndIdx, 
                                         outputTableLocation, outputTableName)
