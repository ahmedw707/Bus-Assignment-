# -*- coding: utf-8 -*-
"""
Created on Wed Apr  3 17:00:47 2019

@author: ahmed
"""

import numpy as np
import pandas as pd
import datetime as dt
import os
from collections import Counter

def bustype(buslist, btype):
    if buslist==[]: return []
    global busprop
    output=[]
    for i in buslist:
        if 'temp' in i: continue
        if busprop[busprop['Bus']==i]['Type'].iloc[0]==btype:
            output.append(i)
    return output

def preference(buslist, ptype):
    if buslist==[]: return []
    global busprop
    output=[]
    for i in buslist:
        if 'temp' in i: continue
        if busprop[busprop['Bus']==i]['Preference'].iloc[0]==ptype:
            output.append(i)
    return output

def AssignedRoute(buslist, A1):
    if buslist==[]: return []
    global busprop
    A2=A1[4:]+"-"+A1[:3]
    output=[]
    for i in buslist:
        if 'temp' in i: continue
        if (busprop[busprop['Bus']==i]['Route'].iloc[0]==A1) or (busprop[busprop['Bus']==i]['Route'].iloc[0]==A2):
            output.append(i)
    return output

def secondary(buslist, routelist):
    if buslist==[]: return []
    global busprop
    output=[]
    
    for i in buslist:
        if 'temp' in i: continue
        if (busprop[busprop['Bus']==i]['Route'].iloc[0] in routelist) or (busprop[busprop['Bus']==i]['Route'].iloc[0] in routelist):
            output.append(i)
    return output

def RouteR(buslist, Dep, Arr):
    if buslist==[]: return []
    global busprop
    output=[]
    for i in buslist:
        if 'temp' in i: continue
        if (Dep not in busprop[busprop['Bus']==i]['Route'].iloc[0]) and (Arr in busprop[busprop['Bus']==i]['Route'].iloc[0]):
            output.append(i)
    return output

def updateR(sbuscount,Rating):
    rank, count, previous, buscountrank = 0, 0, None, {}
    for key, num in sbuscount:    
        count += 0.5
        if num != previous:
            rank += count
            previous = num
            count = 0
        buscountrank[key] = rank
    Ranking = dict(Counter(buscountrank)+Counter(Rating))
    Ranking.update({'temp0': 0,'temp1' : 0, 'temp2' : 0, 'temp3' :0, 'temp4' :0, 'temp5': 0, 'temp6': 0})
    return Ranking

routelist = pd.read_excel('routelist.xlsx', sheet_name='Sheet4')
busprop=pd.read_excel('buses properties.xlsx', sheet_name='Buses3')
Jan=pd.read_excel('book1.xlsx', sheet_name='Sheet3')
stay=pd.read_excel('stay.xlsx', sheet_name='Sheet1')
dtime=pd.read_excel('Date.xlsx', sheet_name='Sheet1')

busprop['Departure']=busprop['Route1'].apply(lambda x:x[:3])
busprop['Arrival']=busprop['Route1'].apply(lambda x:x[4:])
busprop['Route']=busprop["Arrival"] + "-" + busprop["Departure"]

stay['Stay']=stay['Stay'].apply(lambda x:dt.timedelta(hours=x.hour, minutes=x.minute))
mylist=pd.merge(Jan,routelist[['Route','Trip','Km','Step', "Arrival"]],on='Route', how='left')
mylist2=pd.merge(mylist,busprop[['Bus','Year','Breakdown',"Preference"]],on='Bus', how='left')
df=mylist2.drop(columns=["Schedule departure", "Step",'Actual Departure Time','Km','Step','Arrival_y'])
df=df[df["Bus"]!="Drop"]
df=df[df["Bus"]!="DROP"]
df.loc[df['Trip']!=df['Trip'],'Trip']=2359
df["Trip"]=df["Trip"].apply(int)
df["Trip"]=df["Trip"].apply(str)
df["min"]=df["Trip"].apply(lambda x:x[len(x)-2:]).apply(int)
df["hour"]=df["Trip"].apply(lambda x:x[0:len(x)-2] if len(x)>2 else '0').apply(int)
df["triptime"]=df["min"].apply(lambda x: np.timedelta64(x,'m'))+df["hour"].apply(lambda x: np.timedelta64(x,'h'))
df=df.drop(columns=['Trip','min','hour'])

dtime["Dlist"]=dtime["Datetime"].apply(lambda x:x if int(x.strftime("%M"))%15==0 else x+pd.to_timedelta(15-int(x.strftime("%M"))%15,'m'))
df["depdate"]=df["datetime"].apply(lambda x:x if int(x.strftime("%M"))%15==0 else x+pd.to_timedelta(15-int(x.strftime("%M"))%15,'m'))
df["arrdate"]=df["depdate"]+df["triptime"]
df["arrdate"]=df["arrdate"].apply(lambda x:x if int(x.strftime("%M"))%15==0 else x+pd.to_timedelta(15-int(x.strftime("%M"))%15,'m'))
df['triptime']=df['arrdate']-df['depdate']
df['Departure']=df['Route'].apply(lambda x:x[:3])
df['Year'] = df['Year'].apply(lambda x: int(x))
df['Breakdown'] = df['Breakdown'].apply(lambda x: int(x))

gold=list(busprop[busprop['Type']==1]['Bus'])
luxury=list(busprop[busprop['Type']==2]['Bus'])
luxuryr=['LHR-RWP','LHR-MTN','LHR-FSD','FSD-LHR','MTN-LHR','RWP-LHR']
APV=list(busprop[busprop['Type']==4]['Bus'])
APVr=['MRE-RWP','RWP-MRE']
df['Type']=3

for j in gold:
    df.loc[df.Bus==j, 'Type']=1

for j in luxury:
    for i in luxuryr:
        df.loc[(df.Bus==j) & (df.Route11==i), 'Type']=2
      
for j in APV:
    for i in APVr:
        df.loc[(df.Bus==j) & (df.Route11==i), 'Type']=4
        
df=pd.merge(df, stay, on='Route11', how='left')
df.loc[df['Stay']!=df['Stay'],'Stay']=dt.time(0,0)

planfrom=pd.Timestamp('2019-01-22 00:00:00')
plantill=pd.Timestamp('2019-01-29 23:59:00')

dfhalf=df[df['Date']<planfrom]
dfhalf=dfhalf.reset_index()
busloc={}
bustime={}
buscount={}
busAvailability={}
Ranking={}
Rating={}

for i in dfhalf.index:
    busloc[dfhalf.iloc[i]['Bus']]=dfhalf.iloc[i]['Actual Arrival']
    bustime[dfhalf.iloc[i]['Bus']]=dfhalf.iloc[i]['arrdate']
    busAvailability[dfhalf.iloc[i]['Bus']]=dfhalf.iloc[i]['Availability']    

for i in bustime:
    if bustime[i]<planfrom:
        bustime[i]=planfrom-pd.Timedelta('0 days 01:00:00')
          
for i in busprop.index:
    busprop['rate'] = ((busprop.Breakdown.rank(ascending=0,method='min')*0.15))+((busprop.Year.rank(ascending=1,method='min')*0.25))+((busprop.Accident.rank(ascending=0,method='min')*0.1))#+((buscount[i].rank(ascending=1)*0.5))
    Rating[busprop.iloc[i]['Bus']]=busprop.iloc[i]['rate']
#   Ranking = dict(Counter(buscountrank)+Counter(Rating))
    
for i in list(busprop['Bus']):
    buscount[i]=0
    
sbuscount = sorted(buscount.items(), key=lambda item: item[1])
Ranking=updateR(sbuscount,Rating)


tlist=sorted(list(set(list(df['depdate'])+list(dtime['Dlist']))))
tlist=[item for item in tlist if item >=planfrom]
tlist=[planfrom-pd.Timedelta('0 days 01:00:00')]+tlist
cities=list(set(list(df['Actual Arrival'])))

a=df.groupby(['Route11']).count()['Sr']
routeused=pd.DataFrame({'Route':a.index,'Count':a.values})
routeused=routeused[routeused['Count']>14]

rd={}
for city in cities:
    rd[city]={'primaryr':[], 'primaryc':[], 'secondaryr':[]}
    for route in list(routeused['Route']):
        tick=0
        dep=route[:3]
        arr=route[4:]
        if city==dep:
            rd[city]['primaryr'].append(route)
            if arr not in rd[city]['primaryc']:
                pcity=arr
                rd[city]['primaryc'].append(arr)
                tick=1
                
        elif city==arr:
            rd[city]['primaryr'].append(route)
            if dep not in rd[city]['primaryc']:
                pcity=dep
                rd[city]['primaryc'].append(dep)
                tick=1
        if tick==1:
            for route2 in list(set(list(df['Route11']))):
                if pcity in route2 and route2 not in rd[city]['primaryr']:
                    rd[city]['secondaryr'].append(route2)

opsbuses=[]
r={}
for city in cities:
    temp={}
    count=0    
    for t in tlist:
        count+=1 #a vague approximatioon of latest arrival time of previous period
        if count<=200:
            buslist=[]
            for b in busloc:
                if busloc[b]==city and bustime[b]==t and busAvailability[b]=='Available':
                    buslist.append(b)
            temp[t]={'buslist':buslist}
        else:
            temp[t]={'buslist':[]}
        r[city]=temp
               
deponly=df[df['Terminal']==df['Terminal']]
deponly=deponly[deponly['depdate']>=planfrom]
deponly=deponly[deponly['depdate']<plantill]
dfdict=deponly[['Departure','Arrival_x','depdate','triptime','arrdate', 'Type', 'Stay', 'Year', 'Route11', 'Breakdown', 'Preference']].to_dict('split')

count=0
mismatchcount=0
routecount=0
yearmatchcount=0
for t in range(len(tlist)):
    for s in dfdict['data']:
        text=""
#        AvaBuslist=r[s[0]][s[2]]['buslist']
        tick=1
        if s[2]!=tlist[t]:
            tick=0
            continue;
#        AvaBuslist=r[s[0]][s[6]]['buslist']
        if (len(r[s[0]][s[2]]['buslist']))==0:
            if len(r[s[0]][s[2]+pd.Timedelta(minutes=15)]['buslist'])!=0:
                busL=r[s[0]][s[2]+pd.Timedelta(minutes=15)]['buslist']
                bus=busL[0]
                busrating=Ranking[bus]
                for j in busL:
                    if Ranking[j]>busrating:
                        busrating=Ranking[j]
                        bus=j
                r[s[0]][s[2]+pd.Timedelta(minutes=15)]['buslist'].remove(bus)
                text='less stay by 15 mins'
#                AvaBuslist=busL
                
            elif len(r[s[0]][s[2]+pd.Timedelta(minutes=30)]['buslist'])!=0:
                busL=r[s[0]][s[2]+pd.Timedelta(minutes=30)]['buslist']
                bus=busL[0]
                busrating=Ranking[bus]
                for j in busL:
                    if Ranking[j]>busrating:
                        busrating=Ranking[j]
                        bus=j                
                r[s[0]][s[2]+pd.Timedelta(minutes=30)]['buslist'].remove(bus)
                text='less stay by 30 mins'
#                AvaBuslist=busL

            elif (len(r[s[0]][s[2]+pd.Timedelta(minutes=45)]['buslist']))!=0:
                busL=r[s[0]][s[2]+pd.Timedelta(minutes=45)]['buslist']
                bus=busL[0]
                busrating=Ranking[bus]
                for j in busL:
                    if Ranking[j]>busrating:
                        busrating=Ranking[j]
                        bus=j 
                r[s[0]][s[2]+pd.Timedelta(minutes=45)]['buslist'].remove(bus)
                text='less stay by 45 mins'
#                AvaBuslist=busL
               
            else:    
                bus='temp'+str(count)
                count+=1
                text='no bus'  
                
        else:
            usedlist=bustype(r[s[0]][s[2]]['buslist'],s[5])
            Routematch=AssignedRoute(usedlist,s[8]) 
            Arrmatch=RouteR(usedlist,s[0],s[1])
            secondarymatch=secondary(usedlist,rd[s[1]]['secondaryr'])
            
            if len(Routematch)>0:
                busL=Routematch  
                bus=busL[0]
                busrating=Ranking[bus]
                for j in busL:
                    if Ranking[j]>busrating:
                        busrating=Ranking[j]
                        bus=j
                r[s[0]][s[2]]['buslist'].remove(bus)
                text='on route'
#                AvaBuslist=busL
                routecount+=1
                      
            else:
                temptime=s[2]+pd.Timedelta(minutes=15)
                usedlist_temp=bustype(r[s[0]][temptime]['buslist'],s[5])            
                Routematch=AssignedRoute(usedlist_temp,s[8])
                
                if len(Routematch)>0:
                   busL=Routematch
                   bus=busL[0]
                   busrating=Ranking[bus]
                   for j in busL:
                       if Ranking[j]>busrating:
                           busrating=Ranking[j]
                           bus=j                      
                   r[s[0]][temptime]['buslist'].remove(bus)
                   text='on route, 15'
#                   AvaBuslist=busL
                   routecount+=1
                else:
                    temptime=s[2]+pd.Timedelta(minutes=30)
                    usedlist_temp=bustype(r[s[0]][temptime]['buslist'],s[5])
                    Routematch=AssignedRoute(usedlist_temp,s[8])
                    
                    if len(Routematch)>0:
                        busL=Routematch
                        bus=busL[0]
                        busrating=Ranking[bus]
                        for j in busL:
                            if Ranking[j]>busrating:
                                busrating=Ranking[j]
                                bus=j                                              
                        r[s[0]][temptime]['buslist'].remove(bus)
                        text='on route, 30'
#                        AvaBuslist=busL
                        routecount+=1
                        
                    elif len(Arrmatch)>0:
                        busL=Arrmatch
                        bus=busL[0]
                        busrating=Ranking[bus]
                        for j in busL:
                            if Ranking[j]>busrating:
                                busrating=Ranking[j]
                                bus=j      
                        r[s[0]][s[2]]['buslist'].remove(bus)
                        text='Arrival match'
#                        AvaBuslist=busL
                     
                    elif len(secondarymatch)>0:
                        busL=secondarymatch
                        bus=busL[0]
                        busrating=Ranking[bus]
                        for j in busL:
                            if Ranking[j]>busrating:
                                busrating=Ranking[j]
                                bus=j      
                        r[s[0]][s[2]]['buslist'].remove(bus)
                        text='secondary match'
#                        AvaBuslist=busL    

                    elif len(usedlist)>0:
                        busL=preference(usedlist,s[10]) #is 1,0
                        if len(busL)>0:
                            bus=busL[0]
                            busrating=Ranking[bus]
                            for j in busL:
                                if Ranking[j]>busrating:
                                    busrating=Ranking[j]
                                    bus=j                            
                            r[s[0]][s[2]]['buslist'].remove(bus)
                            text='type match'
                        else:
                            bus=usedlist[0]
                            busrating=Ranking[bus]
                            for j in busL:
                                if Ranking[j]>busrating:
                                    busrating=Ranking[j]
                                    bus=j                            
                            r[s[0]][s[2]]['buslist'].remove(bus)
                            text='should avoid'
#                        AvaBuslist=busL
                        
                    else:
                        busL=r[s[0]][s[2]]['buslist']
                        bus=busL[0]
                        busrating=Ranking[bus]
                        for j in busL:
                            if Ranking[j]>busrating:
                                busrating=Ranking[j]
                                bus=j   
                        r[s[0]][s[2]]['buslist'].remove(bus)
                        rt=s[5]
                        bt=busprop[busprop['Bus']==bus]['Type'].iloc[0]
                        text=str(rt)+"-"+str(bt)+"-mismatch"
#                        AvaBuslist=busL
                        mismatchcount+=1

        opsbuses.append(bus)
        s.append(bus)
        s.append(text)
#        s.append(AvaBuslist)
        if bus in buscount.keys():
            buscount[bus]+=1
            sbuscount = sorted(buscount.items(), key=lambda item: item[1])
#        if bus not in Ranking.keys():
#            Ranking[bus]=0
        Ranking = updateR(sbuscount,Rating) 
        route=((s[0]+"-"+s[1]))
        st=stay[stay['Route11']==route]['Stay'].iloc[0]  
        arr=s[4]+st
        if arr not in tlist:
            tlist.append(arr)
            tlist=sorted(tlist)
            for city in cities:
                r[city][arr]={'buslist':[]}                
        r[s[1]][arr]['buslist']=[bus]+r[s[1]][arr]['buslist']
        
       
    if t==len(tlist)-1:
        continue
    for city in cities:
        r[city][tlist[t+1]]['buslist']=r[city][tlist[t]]['buslist']+r[city][tlist[t+1]]['buslist']
        
dfpd=pd.DataFrame.from_dict(dfdict['data'])
dfpd=dfpd.rename(index=str, columns={0:"Dep",1:"Arr",2:"Deptime",3:"traveltime",4:"Arrtime",5:"Type",6:"Stay",7:"Year",8:"Route",9:"Breakdown",10:"Preference",11:"Assignedbus",12:"Text",13:"AvaBuslist"})
dfpd['Assignedbus'].nunique()
Jan['Bus'].nunique()
deponly['Bus'].nunique()
dfpd['Dep_date'] = [d.date() for d in dfpd['Deptime']]
dfpd['Dep_time'] = [d.time() for d in dfpd['Deptime']]
deponly['Dep_date'] = [d.date() for d in deponly['datetime']]
print (dfpd.groupby('Dep_date').Assignedbus.nunique())
print (deponly['Bus'].nunique())
print (dfpd['Assignedbus'].nunique())
print (dfpd['Text'].value_counts())
print (deponly.groupby('Dep_date').Bus.nunique())

usage=pd.DataFrame()
usage['Bus']=list(busprop['Bus'])
for i in dfpd['Dep_date'].unique():
    usage[i]=0
a=dfpd.groupby(['Dep_date','Assignedbus']).count()['Dep']
tempdf=pd.DataFrame({'combination':a.index,'count':a.values})
for i in tempdf.index:
    usage.loc[usage['Bus']==tempdf.iloc[i]['combination'][1],tempdf.iloc[i]['combination'][0]]=tempdf.iloc[i]['count']   