# -*- coding: utf-8 -*-
import os
import glob
import pandas as pd
import math


result=[]

def open_files(way):
    for x in os.walk(way):
        for y in glob.glob(os.path.join(x[0], '*.xlsx')):
            df=pd.read_excel(y)
            #print("1")
            Betreuer=get_Betreuer(y)
            #print("2")
            Einordnung=get_Einordnung(y)
            #print("3")
            Studiengang=get_Studiengang(df)
            Sem=get_Semester(df)
            zwischenliste=[Betreuer, Einordnung]
            zwischenliste.extend(Studiengang)
            zwischenliste.extend(Sem)
            #print(zwischenliste)
            result.append(zwischenliste)
    liste=[]
    liste =sorted(result,key=lambda x: (x[1],x[0]))
    return liste


def get_Betreuer(name):
    """ 'Name' """
    start=name.find(" ")+len("AG")
    end=name.find(".xlsx")
    betreuer=name[start:end]
    return betreuer
    
def get_Einordnung(name):
    """ 'Alte Geschichte' oder 'Mittelalterliche Geschichte' oder 'Neuere Geschichte'"""
    if "Alte Geschichte" in name or "AG" in name:
        einordnung="AG"
    elif "Mittelalterliche Geschichte" in name or "MA" in name:
        einordnung="MA"
    elif "Neuere Geschichte" in name or "NG" in name:
        einordnung="NZ"
    return einordnung

def get_Studiengang(df):
    counter_L2=0
    counter_L3=0
    counter_L5=0
    counter_BA_HF=0
    counter_BA_NF=0
    counter_sonst=0
    counter_unbekannt=0
    s=df["Studiengang"].tolist()
    hf=["BA HF", "Bachelor HF", "BA Hauptfach", "Geschichte Hauptfach","Geschichte HF", "Bachelor Hauptfach", "GE HF", "HF Geschichte"] #Abk können ergänzt werden
    nf= ["BA NF", "Bachelor NF", "BA Nebenfach", "Geschichte Nebenfach", "Geschichte NF", "Bachelor Nebenfach", "GE NF", "NF Geschichte"]
    L2=["L2", "Real", "Haupt"]
    L3=["L3", "Gym"]
    L5=["L5", "Förder"]
    sonst=" " #nicht verändern
    for i in range(len(s)):
        s[i]=str(s[i])
        #weitere Abkürzungen können in der Form "or "xy" in i" eingefügt werden
        if s[i] =="nan":
            counter_unbekannt+=1
        elif any(x in s[i] for x in L2):
            counter_L2+=1
        elif any(x in s[i] for x in L3):
            counter_L3+=1
        elif any(x in s[i] for x in L5):
            counter_L5+=1
        elif any(h in s[i] for h in hf):  
            counter_BA_HF+=1
        elif any(n in s[i] for n in nf):  
            counter_BA_NF+=1
        else:
            counter_sonst+=1
            sonst+=s[i] + ", "
    counter_alle= counter_L2 + counter_L3 + counter_L5 + counter_BA_HF + counter_BA_NF + counter_sonst + counter_unbekannt
    studiengaenge=[counter_alle, counter_L2, counter_L3, counter_L5, counter_BA_HF, counter_BA_NF, counter_sonst, sonst, counter_unbekannt]
    return studiengaenge

def get_Semester(df):
    counter_unbekannt=0
    counter_1=0
    counter_2=0
    counter_3=0
    counter_4=0
    counter_5=0
    counter_6=0
    counter_7=0
    s=df["Fachsemester"].tolist()
    for i in range(len(s)):
        s[i]=str(s[i])
        #weitere Abkürzungen können in der Form "or "xy" in i" eingefügt werden
        if s[i]=="nan":
            counter_unbekannt+=1
        elif s[i] =="1.0":
            counter_1+=1
        elif s[i]=="2.0":
            counter_2+=1
        elif s[i]=="3.0":
            counter_3+=1
        elif s[i]=="4.0":
            counter_4+=1
        elif s[i]=="5.0":  
            counter_5+=1
        elif s[i]=="6.0":  
            counter_6+=1
        else:
            counter_7+=1
    semester=[counter_1 , counter_2 , counter_3 , counter_4 , counter_5 , counter_6 , counter_7, counter_unbekannt]
    return semester


def calc_einordnung(result):
    ag=["Gesamt:", "AG", 0,0,0,0,0,0,0,"", 0, 0,0,0,0,0,0,0,0] #mittleren 8 0 sind für Studiengang, am Ende für Semester
    c_1=0
    ma=["Gesamt:", "MA" , 0,0,0,0,0,0,0,"",0, 0,0,0,0,0,0,0,0]
    c_2=0
    nz=["Gesamt:", "NZ" , 0,0,0,0,0,0,0,"",0, 0,0,0,0,0,0,0,0]
    c_3=0
    for i in range(len(result)):
        if result[i][1]=="AG":
            c_1+=1
            for j in range(2, len(ag)):
                ag[j]+=result[i][j]
        if result[i][1]=="MA":
            c_2+=1
            for j in range(2, len(ma)):
                ma[j]+=result[i][j]
        if result[i][1]=="NZ":
            c_3+=1
            for j in range(2, len(nz)):
                nz[j]+=result[i][j]
    result.insert(c_1, ag)
    result.insert(c_1 + c_2 + 1, ma)
    result.insert(c_1+c_2+c_3 + 2, nz)
    return result
            
            
def make_Output(ausgabepfad, result):
    out=ausgabepfad
    df=pd.DataFrame(result, columns=["Kurs", "Einordnung", "# Studenten", "#L2", "#L3", "#L5", "#BA HF", "#BA NF", "#Sonstige", ":", "#Unbekannt", "#Semester: =1", "=2", "=3", "=4", "=5", "=6", ">6", "Unbekannt"])

    df.to_excel(out + "\Auswertung_Seminar.xlsx", index=0)

def main(eingabepfad, ausgabepfad):                       
    result=open_files(eingabepfad)
    result=calc_einordnung(result)
    make_Output(ausgabepfad, result)


            
#"""Liste=[Kursleiter, Einordnung (AG, MA, NG), [#Studenten, #L2, #L3, #L5,#BA HF, #BA NF, #sonst] #bis 7 Semester, #L2, #L3, #L5, #BA HF, #BA NF, #sonst]"""
# F:\Uni\HiWi\eLearning\PS Auswertung\Ausgang
#Liste für Sonstige
