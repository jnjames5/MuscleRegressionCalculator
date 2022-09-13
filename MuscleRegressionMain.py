# -*- coding: utf-8 -*-
"""
Created on Mon May  9 15:57:56 2022

@author: jjames5
"""
import pandas as pd
import numpy as np

global LevelDict
global CSA_Co
global Dist_Co
global Angle_Co
global MuscleDict
LevelDict={'L4': 13, 'L3':12, 'L2':11, 'L1':10, 'T12':9, 'T11':8, 'T10':7, 'T9':6, 'T8':5, 'T7':4, 'T6':3, 'T5':2, 'T4':1}
#MuscleDict={'PM':"Pectoralis Major", 'RA':"Rectus Abdominis", 'SA':2, 'TR':3, 'LD':4, 'EO':5, 'IO':6, 'PS':7, 'ES':8, 'MU':9, 'QL':10}
CSA_Co=pd.read_excel("S:\Obl\Members\Joanna_James\Muscle_Regression_results_\ExampleTables_05052022.xlsx",'CSA')
CSA_Co.columns=[c.replace(' ',"") for c in CSA_Co.columns]
Dist_Co=pd.read_excel("S:\Obl\Members\Joanna_James\Muscle_Regression_results_\ExampleTables_05052022.xlsx",'Distance')
Dist_Co.columns=[c.replace(' ',"") for c in Dist_Co.columns]
Angle_Co=pd.read_excel("S:\Obl\Members\Joanna_James\Muscle_Regression_results_\ExampleTables_05052022.xlsx",'Angle')
Angle_Co.columns=[c.replace(' ',"") for c in Angle_Co.columns]

#Coefficents: Age, Level, AgexLevel, Height, HeightxLevel, Weight, WeightxLevel, Constant, Level-Specific-Mean
#Muscles:       PM, RA, SA, TR, LD, EO, IO, PS, ES, MU, QL
#Muscle Levels: T9-T4, L4-T10, T11-T6, T10-T4, L3-T4, L4-T10, L4-L2,  L4-L1, L4-T4, L4-T4, L1-L4 

class Muscle:
    def __init__(self, sex, age, height, weight, musc):
        self.sex=sex
        self.age=age
        self.height=height
        self.weight=weight 
        self.musc=musc
    def CSA(self,lvl):
        if self.sex == 'Male':
            self.Age_Const= CSA_Co[(CSA_Co['Sex']=='Men') & (CSA_Co['Muscle']==self.musc)].Age
            self.Level_Const= CSA_Co[(CSA_Co['Sex']=='Men') & (CSA_Co['Muscle']==self.musc)].Level
            self.Age_Level_Const= CSA_Co[(CSA_Co['Sex']=='Men') & (CSA_Co['Muscle']==self.musc)].AgexLevel
            self.Height_Const= CSA_Co[(CSA_Co['Sex']=='Men') & (CSA_Co['Muscle']==self.musc)].Height
            self.Height_Level_Const= CSA_Co[(CSA_Co['Sex']=='Men') & (CSA_Co['Muscle']==self.musc)].HeightxLevel
            self.Weight_Const= CSA_Co[(CSA_Co['Sex']=='Men') & (CSA_Co['Muscle']==self.musc)].Weight
            self.Weight_Level_Const= CSA_Co[(CSA_Co['Sex']=='Men') & (CSA_Co['Muscle']==self.musc)].WeightxLevel
            self.Const= CSA_Co[(CSA_Co['Sex']=='Men') & (CSA_Co['Muscle']==self.musc)].Constant
            self.LSM=CSA_Co[(CSA_Co['Sex']=='Men') & (CSA_Co['Muscle']==self.musc)][lvl]
        elif self.sex=='Female':
            self.Age_Const= CSA_Co[(CSA_Co['Sex']=='Women') & (CSA_Co['Muscle']==self.musc)].Age
            self.Level_Const= CSA_Co[(CSA_Co['Sex']=="Women") & (CSA_Co['Muscle']==self.musc)].Level
            self.Age_Level_Const= CSA_Co[(CSA_Co['Sex']=='Women') & (CSA_Co['Muscle']==self.musc)].AgexLevel
            self.Height_Const= CSA_Co[(CSA_Co['Sex']=='Women') & (CSA_Co['Muscle']==self.musc)].Height
            self.Height_Level_Const= CSA_Co[(CSA_Co['Sex']=='Women') & (CSA_Co['Muscle']==self.musc)].HeightxLevel
            self.Weight_Const= CSA_Co[(CSA_Co['Sex']=='Women') & (CSA_Co['Muscle']==self.musc)].Weight
            self.Weight_Level_Const= CSA_Co[(CSA_Co['Sex']=='Women') & (CSA_Co['Muscle']==self.musc)].WeightxLevel
            self.Const= CSA_Co[(CSA_Co['Sex']=='Women') & (CSA_Co['Muscle']==self.musc)].Constant
            self.LSM=CSA_Co[(CSA_Co['Sex']=='Women') & (CSA_Co['Muscle']==self.musc)][lvl]
            
        csa=self.age*self.Age_Const+self.Level_Const*LevelDict.get(lvl)+self.Age_Level_Const*(self.age*LevelDict.get(lvl))\
            +self.Height_Const*self.height + self.Height_Level_Const*(self.height*LevelDict.get(lvl))\
            +self.Weight_Const*self.weight+self.Weight_Level_Const*(self.weight*LevelDict.get(lvl))\
            + self.Const+self.LSM
        return csa
    def Distance(self,lvl):
        if self.sex == 'Male':
            self.Age_Const= Dist_Co[(Dist_Co['Sex']=='Men') & (Dist_Co['Muscle']==self.musc)].Age
            self.Level_Const= Dist_Co[(Dist_Co['Sex']=='Men') & (Dist_Co['Muscle']==self.musc)].Level
            self.Age_Level_Const= Dist_Co[(Dist_Co['Sex']=='Men') & (Dist_Co['Muscle']==self.musc)].AgexLevel
            self.Height_Const= Dist_Co[(Dist_Co['Sex']=='Men') & (Dist_Co['Muscle']==self.musc)].Height
            self.Height_Level_Const= Dist_Co[(Dist_Co['Sex']=='Men') & (Dist_Co['Muscle']==self.musc)].HeightxLevel
            self.Weight_Const= Dist_Co[(Dist_Co['Sex']=='Men') & (Dist_Co['Muscle']==self.musc)].Weight
            self.Weight_Level_Const= Dist_Co[(Dist_Co['Sex']=='Men') & (Dist_Co['Muscle']==self.musc)].WeightxLevel
            self.Const= Dist_Co[(Dist_Co['Sex']=='Men') & (Dist_Co['Muscle']==self.musc)].Constant
            self.LSM=Dist_Co[(Dist_Co['Sex']=='Men') & (Dist_Co['Muscle']==self.musc)][lvl]
        elif self.sex=='Female':
            self.Age_Const= Dist_Co[(Dist_Co['Sex']=='Women') & (Dist_Co['Muscle']==self.musc)].Age
            self.Level_Const= Dist_Co[(Dist_Co['Sex']=="Women") & (Dist_Co['Muscle']==self.musc)].Level
            self.Age_Level_Const= Dist_Co[(Dist_Co['Sex']=='Women') & (Dist_Co['Muscle']==self.musc)].AgexLevel
            self.Height_Const= Dist_Co[(Dist_Co['Sex']=='Women') & (Dist_Co['Muscle']==self.musc)].Height
            self.Height_Level_Const= Dist_Co[(Dist_Co['Sex']=='Women') & (Dist_Co['Muscle']==self.musc)].HeightxLevel
            self.Weight_Const= Dist_Co[(Dist_Co['Sex']=='Women') & (Dist_Co['Muscle']==self.musc)].Weight
            self.Weight_Level_Const= Dist_Co[(Dist_Co['Sex']=='Women') & (Dist_Co['Muscle']==self.musc)].WeightxLevel
            self.Const= Dist_Co[(Dist_Co['Sex']=='Women') & (Dist_Co['Muscle']==self.musc)].Constant
            self.LSM=Dist_Co[(Dist_Co['Sex']=='Women') & (Dist_Co['Muscle']==self.musc)][lvl]
            
        dist=self.age*self.Age_Const+self.Level_Const*LevelDict.get(lvl)+self.Age_Level_Const*(self.age*LevelDict.get(lvl))\
            +self.Height_Const*self.height + self.Height_Level_Const*(self.height*LevelDict.get(lvl))\
            +self.Weight_Const*self.weight+self.Weight_Level_Const*(self.weight*LevelDict.get(lvl))\
            + self.Const+self.LSM
        return dist
    def Angle(self,lvl):
        if self.sex == 'Male':
            self.Age_Const= Angle_Co[(Angle_Co['Sex']=='Men') & (Angle_Co['Muscle']==self.musc)].Age
            self.Level_Const= Angle_Co[(Angle_Co['Sex']=='Men') & (Angle_Co['Muscle']==self.musc)].Level
            self.Age_Level_Const= Angle_Co[(Angle_Co['Sex']=='Men') & (Angle_Co['Muscle']==self.musc)].AgexLevel
            self.Height_Const= Angle_Co[(Angle_Co['Sex']=='Men') & (Angle_Co['Muscle']==self.musc)].Height
            self.Height_Level_Const= Angle_Co[(Angle_Co['Sex']=='Men') & (Angle_Co['Muscle']==self.musc)].HeightxLevel
            self.Weight_Const= Angle_Co[(Angle_Co['Sex']=='Men') & (Angle_Co['Muscle']==self.musc)].Weight
            self.Weight_Level_Const= Angle_Co[(Angle_Co['Sex']=='Men') & (Angle_Co['Muscle']==self.musc)].WeightxLevel
            self.Const= Angle_Co[(Angle_Co['Sex']=='Men') & (Angle_Co['Muscle']==self.musc)].Constant
            self.LSM=Angle_Co[(Angle_Co['Sex']=='Men') & (Angle_Co['Muscle']==self.musc)][lvl]
        elif self.sex=='Female':
            self.Age_Const= Angle_Co[(Angle_Co['Sex']=='Women') & (Angle_Co['Muscle']==self.musc)].Age
            self.Level_Const= Angle_Co[(Angle_Co['Sex']=="Women") & (Angle_Co['Muscle']==self.musc)].Level
            self.Age_Level_Const= Angle_Co[(Angle_Co['Sex']=='Women') & (Angle_Co['Muscle']==self.musc)].AgexLevel
            self.Height_Const= Angle_Co[(Angle_Co['Sex']=='Women') & (Angle_Co['Muscle']==self.musc)].Height
            self.Height_Level_Const= Angle_Co[(Angle_Co['Sex']=='Women') & (Angle_Co['Muscle']==self.musc)].HeightxLevel
            self.Weight_Const= Angle_Co[(Angle_Co['Sex']=='Women') & (Angle_Co['Muscle']==self.musc)].Weight
            self.Weight_Level_Const= Angle_Co[(Angle_Co['Sex']=='Women') & (Angle_Co['Muscle']==self.musc)].WeightxLevel
            self.Const= Angle_Co[(Angle_Co['Sex']=='Women') & (Angle_Co['Muscle']==self.musc)].Constant
            self.LSM=Angle_Co[(Angle_Co['Sex']=='Women') & (Angle_Co['Muscle']==self.musc)][lvl]
            
        angle=self.age*self.Age_Const+self.Level_Const*LevelDict.get(lvl)+self.Age_Level_Const*(self.age*LevelDict.get(lvl))\
            +self.Height_Const*self.height + self.Height_Level_Const*(self.height*LevelDict.get(lvl))\
            +self.Weight_Const*self.weight+self.Weight_Level_Const*(self.weight*LevelDict.get(lvl))\
            + self.Const+self.LSM
        return angle
    
#Main Functions for creating muscle, running regression and writing to excel
   
def CreateMuscle(Sex,Age,Height,Weight,Type):
    M=Muscle(Sex,Age,Height,Weight,Type)
    return M

def Generate_Results(M, Levels):
    CSA_Output=[]
    Dist_Output=[]
    Ang_Output=[]
    for j in Levels:
        CSA_Output.append(M.CSA(j).values[0])
        Dist_Output.append(M.Distance(j).values[0])
        Ang_Output.append(M.Angle(j).values[0])
    CSA_Output_All.append(CSA_Output)
    Dist_Output_All.append(Dist_Output)
    Ang_Output_All.append(Ang_Output)
    
    return CSA_Output_All,Dist_Output_All,Ang_Output_All

# Function takes three lists, will convert to dataframe before export to excel
def Write_Results(CSA_Output_All,Dist_Output_All,Ang_Output_All, All_Levels,All_Muscles):
    
    Output_CSA=np.asarray(CSA_Output_All)
    Output_CSA=pd.DataFrame(Output_CSA,columns=All_Levels,index=All_Muscles)

    Output_Dist=np.asarray(Dist_Output_All)
    Output_Dist=pd.DataFrame(Output_Dist,columns=All_Levels,index=All_Muscles)

    Output_Ang=np.asarray(Ang_Output_All)
    Output_Ang=pd.DataFrame(Output_Ang,columns=All_Levels,index=All_Muscles)
    
    with pd.ExcelWriter("Subject_113.xlsx") as writer:
        Output_CSA.to_excel(writer,sheet_name="CSA_Results")
        Output_Dist.to_excel(writer,sheet_name="Distance_Results")
        Output_Ang.to_excel(writer,sheet_name="Angle_Results")


##Example Main

#All Muscles, All Appropriate Levels
All_Muscles=["PM", "RA", "SA","TR", "LD", "EO", "IO", "PS", "ES", "MU", "QL"]       
All_Levels=['L4', 'L3', 'L2', 'L1', 'T12', 'T11', 'T10', 'T9', 'T8', 'T7', 'T6', 'T5', 'T4']     
CSA_Output_All=[]
Dist_Output_All=[]
Ang_Output_All=[]
for i in All_Muscles:
    a=CreateMuscle("Male",72,156.0,58.0,i)
    CSA_Output_All,Dist_Output_All,Ang_Output_All=Generate_Results(a, All_Levels)
Write_Results(CSA_Output_All, Dist_Output_All, Ang_Output_All, All_Levels,All_Muscles)
