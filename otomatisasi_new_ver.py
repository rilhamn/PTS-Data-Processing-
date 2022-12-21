from tkinter import *
from tkinter import ttk
from pandastable import Table, TableModel
from matplotlib.backends.backend_tkagg import (FigureCanvasTkAgg,
NavigationToolbar2Tk)

import pandas as pd
from pandas import ExcelWriter
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
import numpy as np
import seaborn as sns
from statistics import mean
from scipy.stats import linregress
import math
import re

import os
path="D:\Pribadi\Materi Belajar\Materi Magang\PT Medco Power Indonesia\modelling\ijen modelling\Python Testing"
os.chdir(path)
current_directory = os.getcwd()
#======================================== FUNCTION PROCESSING ===========================================#
def process():
    global data_base
    data_base={"geometry":data_base["geometry"]}

    raw_data=pd.ExcelFile(input_process.get()+".xlsx")
    list_log=[]

    for i in raw_data.sheet_names:
        if i[0:2]=="LD" or i[0:2]=="LU":
            list_log.append(i)

    for i in list_log:
        data_base[i]=pivot_table_avg(clean(raw_data.parse(i)))

    # Slope
    data_base["slope"]=slope(data_base)
    label_slope_value=Label(tab_2,text=f"(mean={round(data_base['slope'].loc[:,'SLOPE'].mean(),2)} and median={round(data_base['slope'].loc[:,'SLOPE'].median(),2)})")
    label_slope_value.grid(row=1,column=0,columnspan=3,padx=3,pady=2,sticky=W)
    label_slope_value.config(font=("HP Simplified",10),background="white")

    # avg
    list_minimum=[]
    list_maximum=[]
    for k,v in data_base.items():
        if k!="avg" and k!="geometry" and k!="slope":
            list_minimum.append(v.iloc[0,0])
            list_maximum.append(v.iloc[-1,0])
    minimum=np.min(list_minimum)
    maximum=np.max(list_maximum)

    data_base["avg"]=pd.DataFrame(columns=["DEPTH","PRESSURE","TEMP EXT","FLUID VELOCITY","MASS RATE"])
    data_base["avg"]["DEPTH"]=[i for i in range(minimum,maximum+2,2)]

    string_log=", ".join(list_log)
    label_list_log=Label(tab_3,text=f"Log: {string_log}",background="white",font=("HP Simplified",10))
    label_list_log.grid(row=1,column=0,columnspan=3,padx=3,pady=2,sticky=W)
    label_list_log.config(font=("HP Simplified",10),background="white")

def pivot_table_avg(df):
    df=df.sort_values("DEPTH")
    if math.floor(df.iloc[0,0])%2==0:
        minimum=math.floor(df.iloc[0,0])-1
    else:
        minimum=math.floor(df.iloc[0,0])
    if math.ceil(df.iloc[-1,0])%2==0:
        maximum=math.ceil(df.iloc[-1,0])+1
    else:
        maximum=math.ceil(df.iloc[-1,0])
    df["VELOCITY"]=df["VELOCITY"].apply(lambda x: x/60)
    df_pivot=df[["PRESSURE","TEMP EXT","SPINNER","VELOCITY"]].groupby(pd.cut(df.DEPTH, [i for i in range(minimum,maximum+2,2)])).mean().rename_axis("DEPTH RANGE").reset_index()
    df_pivot["DEPTH RANGE"]=df_pivot["DEPTH RANGE"].apply(lambda x: str(x))
    df_pivot.insert(0,"DEPTH",df_pivot["DEPTH RANGE"].str.findall("(?:-|)\d+").apply(lambda x: int((int(x[0])+int(x[1]))/2)))
    # df_pivot.insert(0,"DEPTH",[i for i in range(minimum+1,maximum+1,2)])

    return df_pivot

def clean(df):
    df=df.dropna(axis=0,how="all")
    df.columns=[i.upper() for i in list(df.columns)]
    df=df.loc[:,["DEPTH","PRESSURE","TEMP EXT","VELOCITY","SPINNER"]].astype(float)
    return  df

#======================================== FUNCTION SHOW GRAPH ==========================================#
#Function Well Trajectory
def plot_well_trajectory():
    fig = Figure(figsize=(6,4.2))
    ax = fig.add_subplot(111)
    list_well_trajectory_depth=data_base["geometry"].copy().loc[:,"DEPTH"]*-1
    list_well_trajectory_id=data_base["geometry"].copy().loc[:,"ID"]
    list_well_trajectory_inc=data_base["geometry"].copy().loc[:,"INCLINATION"]
    list_well_trajectory_type=data_base["geometry"].copy().loc[:,"TYPE"]
    color={"casing":"royalblue","liner":"green","open hole":"brown"}
    line={"casing":"solid","liner":"dashed","open hole":"dashed"}

    loop=0
    for i,j,k,l in zip(list_well_trajectory_depth,list_well_trajectory_id,list_well_trajectory_inc,list_well_trajectory_type):
        if loop==0:
            ax.plot([float(j),float(j)],[0,i],color=color[l],linestyle=line[l])
            ax.plot([-1*float(j),-1*float(j)],[0,i],color=color[l],linestyle=line[l])
            ax.text(float(j)+0.01,i/2,f"Inc:{k}\u00b0")
        else:
            ax.plot([float(j),float(j)],[list_well_trajectory_depth[loop-1],i],color=color[l],linestyle=line[l])
            ax.plot([-1*float(j),-1*float(j)],[list_well_trajectory_depth[loop-1],i],color=color[l],linestyle=line[l])
            ax.text(float(j)+0.01,((i-list_well_trajectory_depth[loop-1])/2)+list_well_trajectory_depth[loop-1],f"Inc:{k}\u00b0")
        loop=loop+1
    ax.set_title(f"Well Trajectory Profile",size = 14)
    ax.set_ylabel("Measured Depth (m)")
    ax.set_xlabel("ID (m)")
    ax.set_xlim(-0.6,0.6)
    ax.grid(axis ="both")

    canvas=FigureCanvasTkAgg(fig,label_showgraph_tab_1)
    canvas.draw()
    canvas.get_tk_widget().grid(row=1,column=0,rowspan=30,columnspan=1000,padx=10,pady=10,sticky=N)

    toolbar=NavigationToolbar2Tk(canvas,label_showgraph_tab_1,pack_toolbar=False)
    toolbar.update()
    toolbar.grid(row=40,column=0,rowspan=1,columnspan=1000,padx=10,pady=10,sticky=W)

#Function Plot Properties all log
def plot_properties_all_log():
    dict_selected_data={}
    for k,v in data_base.items():
        if k!="avg" and k!="geometry" and k!="slope":
            dict_selected_data[f"{k}"]=v.copy()
            dict_selected_data[f"{k}"].loc[:,"DEPTH"]=v.copy().loc[:,"DEPTH"].apply(lambda x: x*-1)
    properties=var_graph_tab_2.get().upper()

    fig = Figure(figsize=(6,4.2))
    ax = fig.add_subplot(111)
    legend=[]
    for k,v in dict_selected_data.items():
        if properties=="TEMP EXT" and k=="avg":
            v.plot(x="BPD",y="DEPTH",ax=ax)
            legend.append("BPD")
        v.plot(x=properties,y="DEPTH",ax=ax)
        legend.append(k)
    ax.set_title(f"{properties} Profile",size = 14)
    ax.set_ylabel("Depth")
    ax.set_xlabel(f"{properties}")
    if properties=="MASS RATE":
        ax.set_xlim(0,100)
    ax.legend(legend,loc="upper right")
    ax.grid(axis ="both")

    canvas=FigureCanvasTkAgg(fig,label_showgraph_tab_2)
    canvas.draw()
    canvas.get_tk_widget().grid(row=1,column=0,rowspan=30,columnspan=1000,padx=10,pady=10,sticky=N)

    toolbar=NavigationToolbar2Tk(canvas,label_showgraph_tab_2,pack_toolbar=False)
    toolbar.update()
    toolbar.grid(row=40,column=0,rowspan=1,columnspan=1000,padx=10,pady=10,sticky=W)

#Function Plot Properties
def plot_properties_avg_log():
    dict_selected_data={}
    dict_selected_data["avg"]=data_base["avg"].copy()
    dict_selected_data["avg"].loc[:,"DEPTH"]=data_base["avg"].copy().loc[:,"DEPTH"].apply(lambda x: x*-1)
    properties=var_graph_tab_3.get().upper()

    fig = Figure(figsize=(6,4.2))
    ax = fig.add_subplot(111)
    legend=[]
    for k,v in dict_selected_data.items():
        if properties=="TEMP EXT" and k=="avg":
            v.plot(x="BPD",y="DEPTH",ax=ax)
            legend.append("BPD")
        v.plot(x=properties,y="DEPTH",ax=ax)
        legend.append(k)
    ax.set_title(f"{properties} Profile",size = 14)
    ax.set_ylabel("Depth")
    ax.set_xlabel(f"{properties}")
    if properties=="MASS RATE":
        ax.set_xlim(-100,100)
    ax.legend(legend,loc="upper right")
    ax.grid(axis ="both")

    canvas=FigureCanvasTkAgg(fig,label_showgraph_tab_3)
    canvas.draw()
    canvas.get_tk_widget().grid(row=1,column=0,rowspan=30,columnspan=1000,padx=10,pady=10,sticky=N)

    toolbar=NavigationToolbar2Tk(canvas,label_showgraph_tab_3,pack_toolbar=False)
    toolbar.update()
    toolbar.grid(row=40,column=0,rowspan=1,columnspan=1000,padx=10,pady=10,sticky=W)

#================================= FUNCTION SLOPE AND FLUID VELOCITY ================================#
def slope(data_base):
    list_minimum=[]
    list_maximum=[]

    for k,v in data_base.items():
        if k!="avg" and k!="geometry" and k!="slope":
            list_minimum.append(v.iloc[0,0])
            list_maximum.append(v.iloc[-1,0])

    minimum=np.min(list_minimum)
    maximum=np.max(list_maximum)

    list_slope=[]
    list_depth=[]
    for i in range(minimum,maximum+2,2):
        x=[]
        y=[]
        for k,v in data_base.items():
            if k!="avg" and k!="geometry" and k!="slope":
                v_temp=v.copy().set_index("DEPTH")
                if i in v_temp.index:
                    x.append(v_temp.loc[i,"SPINNER"])
                    y.append(v_temp.loc[i,"VELOCITY"])
                else:
                    x.append(np.NaN)
                    y.append(np.NaN)
        if len(x)>=2 and len(y)>=2:
            slope=linregress(y,x)[0]
            list_slope.append(slope)
        else:
            list_slope.append(np.NaN)
        list_depth.append(i)

    df_slope=pd.DataFrame(list(zip(list_depth,list_slope)),columns =["DEPTH","SLOPE"])
    return df_slope

def calculate_velocity():
    selected_slope=float(input_slope.get())
    for k,v in data_base.items():
        if k!="slope" and k!="geometry" and k!="avg":
            v["FLUID VELOCITY"]=v[["SPINNER","VELOCITY"]].apply(lambda x: (x["SPINNER"]/selected_slope)-x["VELOCITY"],axis=1)

#======================================== FUNCTION AVERAGE =========================================#
def calculate_average():
    selected_data_pressure=input_average_log_pressure.get().split(",")
    selected_data_temperature=input_average_log_temperature.get().split(",")
    selected_data_fluid_velocity=input_average_log_fluid_velocity.get().split(",")

    dict_temp={
        "PRESSURE":selected_data_pressure,
        "TEMP EXT":selected_data_temperature,
        "FLUID VELOCITY":selected_data_fluid_velocity
    }

    data_base["avg"]=data_base["avg"].set_index("DEPTH")

    for k,v in dict_temp.items():
        data_base["avg"].loc[:,k]=np.NaN
        df_avg_temp=pd.DataFrame(columns=["DEPTH",k])
        for log in v:
            try:
                df_avg_temp=pd.concat([df_avg_temp,data_base[log].loc[:,["DEPTH",k]]])
            except:
                pass
        df_avg_temp=df_avg_temp.groupby("DEPTH").mean()
        data_base["avg"].update(df_avg_temp)
    data_base["avg"]=data_base["avg"].reset_index()

    if "BPD" not in data_base["avg"].columns:
        data_base["avg"].insert(3,"BPD",data_base["avg"]["PRESSURE"].apply(lambda x: (-1.60646649*10**(-4)*(math.log(x))**6)+(8.27534137*10**(-4)*(math.log(x))**5)+(2.43995895*10**(-2)*(math.log(x))**4)+(0.22273639*(math.log(x))**3)+(2.35059677*(math.log(x))**2)+(27.89734893*(math.log(x)))+99.66703174))

    calculate_mass_rate()
#====================================== FUNCTION MASSRATE =========================================#
def add_well():
    if int(input_well_depth.get()) in data_base["geometry"]["DEPTH"].unique():
        data_base["geometry"].loc[data_base["geometry"][data_base["geometry"]["DEPTH"]==int(input_well_depth.get())].index[0],"ID"]=input_well_diameter.get()
        data_base["geometry"].loc[data_base["geometry"][data_base["geometry"]["DEPTH"]==int(input_well_depth.get())].index[0],"INCLINATION"]=input_well_inclination.get()
        data_base["geometry"].loc[data_base["geometry"][data_base["geometry"]["DEPTH"]==int(input_well_depth.get())].index[0],"TYPE"]=var_well_type.get()
    else:
        df_temp=pd.DataFrame({"DEPTH":[int(input_well_depth.get())],"ID":[input_well_diameter.get()],"INCLINATION":[input_well_inclination.get()],"TYPE":[var_well_type.get()]})
        data_base["geometry"]=pd.concat([data_base["geometry"],df_temp])
        data_base["geometry"]=data_base["geometry"].reset_index(drop=True)
        data_base["geometry"]=data_base["geometry"].sort_values("DEPTH",ignore_index=True)       

def remove_well():
    if int(input_well_depth.get()) in data_base["geometry"]["DEPTH"].unique():
        data_base["geometry"]=data_base["geometry"].drop(index=data_base["geometry"][data_base["geometry"]["DEPTH"]==int(input_well_depth.get())].index[0])

def calculate_mass_rate():
    # empty
    data_base["avg"][["ID","AREA","INCLINATION","dp/dz","dv/dz","GRAV TERM","FRIC TERM","ACC TERM","RHO MIX","MASS RATE"]]=np.NaN

    # set well geometry
    data_base["geometry"]=data_base["geometry"].sort_values("DEPTH",ascending=TRUE).reset_index(drop=True)
    data_base["avg"]=data_base["avg"].set_index("DEPTH")

    loop=0
    for i,j,k in zip(data_base["geometry"]["DEPTH"],data_base["geometry"]["ID"],data_base["geometry"]["INCLINATION"]):
        if loop==0:
            data_base["avg"].loc[data_base["avg"].index[0]:i+1,"ID"]=float(j)
            data_base["avg"].loc[data_base["avg"].index[0]:i+1,"INCLINATION"]=float(k)
        else:
            data_base["avg"].loc[i_previous:i+1,"ID"]=float(j)
            data_base["avg"].loc[i_previous:i+1,"INCLINATION"]=float(k)
        i_previous=i
        loop=loop+1

    # Area
    data_base["avg"]["AREA"]=data_base["avg"]["ID"].apply(lambda x: 0.25*math.pi*(x**2))

    # dp/dz and dv/dz
    loop=0
    for i in data_base["avg"].index:
        if loop==0:
            data_base["avg"].loc[i,"dp/dz"]=np.NaN
            data_base["avg"].loc[i,"dv/dz"]=np.NaN
        else:
            data_base["avg"].loc[i,"dp/dz"]=((data_base["avg"].loc[i,"PRESSURE"]-data_base["avg"].loc[i-2,"PRESSURE"])/(2))*100000
            data_base["avg"].loc[i,"dv/dz"]=((data_base["avg"].loc[i,"FLUID VELOCITY"]-data_base["avg"].loc[i-2,"FLUID VELOCITY"])/(2))
        loop=loop+1

    # gravitational term
    data_base["avg"]["GRAV TERM"]=9.81
    data_base["avg"]["GRAV TERM"]=data_base["avg"][["GRAV TERM","INCLINATION"]].apply(lambda x: x["GRAV TERM"]*math.acos(math.radians(x["INCLINATION"])),axis=1)

    # fric term
    data_base["avg"]["FRIC TERM"]=data_base["avg"][["ID","FLUID VELOCITY"]].apply(lambda x:(0.01*x["FLUID VELOCITY"]**2)/(2*x["ID"]),axis=1)

    # acc term
    data_base["avg"]["ACC TERM"]=data_base["avg"][["dv/dz","FLUID VELOCITY"]].apply(lambda x:x["dv/dz"]*x["FLUID VELOCITY"],axis=1)

    # rho mix
    if density_fluid_type.get()=="mix fluid":
        data_base["avg"]["RHO MIX"]=data_base["avg"][["dp/dz","GRAV TERM","FRIC TERM","ACC TERM"]].apply(lambda x:x["dp/dz"]/(x["GRAV TERM"]+x["FRIC TERM"]+x["ACC TERM"]),axis=1)
    elif density_fluid_type.get()=="all steam":
        data_base["avg"]["RHO MIX"]=data_base["avg"]["TEMP EXT"].apply(lambda x: math.exp((1.14E-14*(x**6))-(9.70315E-12*(x**5))+(2.45004487E-09*(x**4))+(0.000000164107501*(x**3))-(0.000221879417*(x**2))+(0.0669727562*x)-5.31406813))
    elif density_fluid_type.get()=="all liquid":
        data_base["avg"]["RHO MIX"]=data_base["avg"]["TEMP EXT"].apply(lambda x: (-2.23821695E-28*(x**12 ))+(4.94472419E-23*(x**10))-(1.48693151E-18*(x**8))-(4.68969075E-13*(x**6))+(0.000000047662926*(x**4))-(4.5273543928009E-03*(x**2))+999.39961079)

    # mass rate
    data_base["avg"]["MASS RATE"]=data_base["avg"][["RHO MIX","FLUID VELOCITY","AREA"]].apply(lambda x:x["RHO MIX"]*x["FLUID VELOCITY"]*x["AREA"],axis=1)

    # reset
    data_base["avg"]=data_base["avg"].reset_index()

#====================================== FUNCTION SAVE FILE =========================================#
def save_xlsx():
    with pd.ExcelWriter(f"{input_save.get()}.xlsx") as writer:
        for k, df in data_base.items():
            df.to_excel(writer,f"{k}")

#================================ INTIAL CONDITION AND PROCESS =====================================#
df_avg=pd.DataFrame(columns=["DEPTH","PRESSURE","TEMP EXT","SPINNER","VELOCITY","SLOPE","MASS RATE"])
data_base={}
data_base["geometry"]=pd.DataFrame(columns=["DEPTH","ID","INCLINATION"])

all_input_diameter=[]
all_input_depth=[]

root=Tk()
root.title("PTS Log Raw Data Processing")
root.config(background="white")
root.geometry("900x660")
root.resizable(False, False)

label_creator=Label(root,text="Created by Ilham Narendrodhipo as part of internship project in Medco Cahaya Geothermal")
label_creator.grid(row=500,column=0,columnspan=3,padx=3,pady=2,sticky=SW)
label_creator.config(font=("HP Simplified",8),background="white")

# Tab
header_style = ttk.Style()
header_style.theme_create( "headerstyle", parent="alt", settings={
        "TNotebook": {"configure": {"tabmargins": [2, 5, 2, 0],"background":"white"}},
        "TNotebook.Tab": {"configure": {"padding": [20,10],"background":"white","font":["HP Simplified",12]}}})
header_style.theme_use("headerstyle")
notebook = ttk.Notebook(root,width=885,height=580)

page_style= ttk.Style()
page_style.configure("pagestyle.TFrame", background="white")
tab_1 = ttk.Frame(notebook,style="pagestyle.TFrame")
tab_2 = ttk.Frame(notebook,style="pagestyle.TFrame")
tab_3 = ttk.Frame(notebook,style="pagestyle.TFrame")
tab_4 = ttk.Frame(notebook,style="pagestyle.TFrame")

notebook.add(tab_1,text="A. Intial Preparation")
notebook.add(tab_2,text="B. Slope Calibarition")
notebook.add(tab_3, text="C. Averaging and Calculate Massrate")
notebook.add(tab_4, text="D. Save Processed Data")
notebook.grid(row=1,column=0,columnspan=100,rowspan=20,padx=3,pady=2,sticky=NW)

## Tab_1 GUI
###
label_process_title=Label(tab_1,text="Load Raw Data")
label_process_title.grid(row=0,column=0,columnspan=3,padx=3,pady=2,sticky=W)
label_process_title.config(font=("HP Simplified bold",12),background="white")
label_process=Label(tab_1,text="File Name:                  ")
label_process.grid(row=1,column=0,columnspan=1,padx=3,pady=2,sticky=W)
label_process.config(font=("HP Simplified",11),background="white")
input_process=Entry(tab_1,width=16,relief=SOLID)
input_process.grid(row=1,column=1,columnspan=2,padx=3,pady=2,sticky=E)
input_process.config(font=("HP Simplified",10),background="white")
button_excute_process = Button(tab_1,text="Execute",command=process)
button_excute_process.grid(row=2,column=2,columnspan=1,padx=3,pady=2,sticky=E)
button_excute_process.config(font=("HP Simplified",9),background="white")
###
spacer1 = Label(tab_1, text="",background="white")
spacer1.grid(row=3, column=0)
label_mass_rate_title=Label(tab_1,text="Well Trajectory")
label_mass_rate_title.grid(row=4,column=0,columnspan=5,padx=3,pady=2,sticky=W)
label_mass_rate_title.config(font=("HP Simplified bold",12),background="white")
label_well_depth=Label(tab_1,text="Depth (m):")
label_well_depth.grid(row=5,column=0,columnspan=1,padx=3,pady=2,sticky=W)
label_well_depth.config(font=("HP Simplified",11),background="white")
input_well_depth=Entry(tab_1,width=16,relief=SOLID)
input_well_depth.grid(row=5,column=1,columnspan=2,padx=3,pady=2,sticky=E)
input_well_depth.config(font=("HP Simplified",10),background="white")
label_well_diameter=Label(tab_1,text="ID (m):")
label_well_diameter.grid(row=6,column=0,columnspan=1,padx=3,pady=2,sticky=W)
label_well_diameter.config(font=("HP Simplified",11),background="white")
input_well_diameter=Entry(tab_1,width=16,relief=SOLID)
input_well_diameter.grid(row=6,column=1,columnspan=2,padx=3,pady=2,sticky=E)
input_well_diameter.config(font=("HP Simplified",10),background="white")
label_well_inclination=Label(tab_1,text=f"Inclination (\u00b0):")
label_well_inclination.grid(row=7,column=0,columnspan=1,padx=3,pady=2,sticky=W)
label_well_inclination.config(font=("HP Simplified",11),background="white")
input_well_inclination=Entry(tab_1,width=16,relief=SOLID)
input_well_inclination.grid(row=7,column=1,columnspan=2,padx=3,pady=2,sticky=E)
input_well_inclination.config(font=("HP Simplified",10),background="white")
label_well_type=Label(tab_1,text="Well Type:")
label_well_type.grid(row=8,column=0,columnspan=1,padx=3,pady=2,sticky=W)
label_well_type.config(font=("HP Simplified",11),background="white")
var_well_type=StringVar()
var_well_type.set("casing")
input_well_type=OptionMenu(tab_1,var_well_type,"casing","liner","open hole")
input_well_type.grid(row=8,column=1,columnspan=3,padx=3,pady=2,sticky=NE)
input_well_type.config(font=("HP Simplified",9),background="white")
add_well_button = Button(tab_1, text="Add", command=add_well)
add_well_button.grid(row=9,column=1,columnspan=1,padx=3,pady=2,sticky=E)
add_well_button.config(font=("HP Simplified",9),background="white")
remove_well_button = Button(tab_1, text="Remove", command=remove_well)
remove_well_button.grid(row=9,column=2,columnspan=1,padx=3,pady=2,sticky=E)
remove_well_button.config(font=("HP Simplified",9),background="white")
###
spacer1 = Label(tab_1, text="",background="white")
spacer1.grid(row=0,column=4)
label_showgraph_tab_1=LabelFrame(tab_1,text="Show Well Trajectory",height=570,width=620)
label_showgraph_tab_1.grid(row=0,column=5,rowspan=1000,columnspan=1000,padx=3,pady=2,sticky=NW)
label_showgraph_tab_1.config(font=("HP Simplified bold",12),background="white")
###
button_excute_showgraph_tab_1=Button(label_showgraph_tab_1,text="Refresh",command=plot_well_trajectory)
button_excute_showgraph_tab_1.grid(row=0,column=0,columnspan=1,padx=3,pady=2,sticky=NW)
label_showgraph_tab_1.grid_propagate(False)
button_excute_showgraph_tab_1.config(font=("HP Simplified",9),background="white")

## Tab_2 GUI
###
label_slope_title=Label(tab_2,text="Calculate Fluid Velocity")
label_slope_title.grid(row=0,column=0,columnspan=3,padx=3,pady=2,sticky=W)
label_slope_title.config(font=("HP Simplified bold",12),background="white")
spacer1 = Label(tab_2, text="",background="white")
spacer1.grid(row=1, column=0)
label_slope=Label(tab_2,text="Slope:                          ")
label_slope.grid(row=2,column=0,columnspan=1,padx=3,pady=2,sticky=W)
label_slope.config(font=("HP Simplified",11),background="white")
input_slope=Entry(tab_2,width=16,relief=SOLID)
input_slope.grid(row=2,column=1,columnspan=2,padx=3,pady=2,sticky=E)
input_slope.config(font=("HP Simplified",10),background="white")
button_excute_velocity = Button(tab_2,text="Execute",command=calculate_velocity)
button_excute_velocity.grid(row=3,column=2,columnspan=1,padx=3,pady=2,sticky=E)
button_excute_velocity.config(font=("HP Simplified",9),background="white")
###
spacer1 = Label(tab_2, text="",background="white")
spacer1.grid(row=0,column=4)
label_showgraph_tab_2=LabelFrame(tab_2,text="Show Properties of All Log",height=570,width=620)
label_showgraph_tab_2.grid(row=0,column=5,rowspan=1000,columnspan=1000,padx=3,pady=2,sticky=NW)
label_showgraph_tab_2.grid_propagate(False)
label_showgraph_tab_2.config(font=("HP Simplified bold",12),background="white")
###
var_graph_tab_2=StringVar()
var_graph_tab_2.set("pressure")
input_properties_tab_2=OptionMenu(label_showgraph_tab_2,var_graph_tab_2,"pressure","temp ext","spinner","velocity","fluid velocity")
input_properties_tab_2.grid(row=0,column=5,columnspan=1,padx=3,pady=2,sticky=NW)
input_properties_tab_2.config(font=("HP Simplified",9),background="white")
button_excute_showgraph_tab_2=Button(label_showgraph_tab_2,text="Show Graph",command=plot_properties_all_log)
button_excute_showgraph_tab_2.grid(row=0,column=6,columnspan=1,padx=3,pady=2,sticky=NW)
button_excute_showgraph_tab_2.config(font=("HP Simplified",9),background="white")

## Average GUI
label_average_title=Label(tab_3,text="Averaging and Calculate Massrate")
label_average_title.grid(row=0,column=0,columnspan=3,padx=3,pady=2,sticky=W)
label_average_title.config(font=("HP Simplified bold",12),background="white")
spacer1 = Label(tab_3, text="",background="white")
spacer1.grid(row=1,column=0)
label_average_log_pressure=Label(tab_3,text="Log (pressure):  ")
label_average_log_pressure.grid(row=2,column=0,columnspan=1,padx=3,pady=2,sticky=W)
label_average_log_pressure.config(font=("HP Simplified",11),background="white")
input_average_log_pressure=Entry(tab_3,width=16,relief=SOLID)
input_average_log_pressure.grid(row=2,column=1,columnspan=2,padx=3,pady=2,sticky=E)
input_average_log_pressure.config(font=("HP Simplified",10),background="white")
label_average_log_temperature=Label(tab_3,text="Log (temperature):")
label_average_log_temperature.grid(row=3,column=0,columnspan=1,padx=3,pady=2,sticky=W)
label_average_log_temperature.config(font=("HP Simplified",11),background="white")
input_average_log_temperature=Entry(tab_3,width=16,relief=SOLID)
input_average_log_temperature.grid(row=3,column=1,columnspan=2,padx=3,pady=2,sticky=E)
input_average_log_temperature.config(font=("HP Simplified",10),background="white")
label_average_log_fluid_velocity=Label(tab_3,text="Log (fluid vel.):  ")
label_average_log_fluid_velocity.grid(row=4,column=0,columnspan=1,padx=3,pady=2,sticky=W)
label_average_log_fluid_velocity.config(font=("HP Simplified",11),background="white")
input_average_log_fluid_velocity=Entry(tab_3,width=16,relief=SOLID)
input_average_log_fluid_velocity.grid(row=4,column=1,columnspan=2,padx=3,pady=2,sticky=E)
input_average_log_fluid_velocity.config(font=("HP Simplified",10),background="white")
label_fluid_type=Label(tab_3,text="Fluid Type:  ")
label_fluid_type.grid(row=5,column=0,columnspan=1,padx=3,pady=2,sticky=W)
label_fluid_type.config(font=("HP Simplified",11),background="white")
density_fluid_type=StringVar()
density_fluid_type.set("mix fluid")
input_density_fluid_type=OptionMenu(tab_3,density_fluid_type,"mix fluid","all steam","all liquid")
input_density_fluid_type.grid(row=5,column=2,columnspan=1,padx=3,pady=2,sticky=E)
input_density_fluid_type.config(font=("HP Simplified",9),background="white")
button_excute_average = Button(tab_3,text="Execute",command=calculate_average)
button_excute_average.grid(row=6,column=2,columnspan=1,padx=3,pady=2,sticky=E)
button_excute_average.config(font=("HP Simplified",9),background="white")
###
spacer1 = Label(tab_3, text="",background="white")
spacer1.grid(row=0,column=4)
label_showgraph_tab_3=LabelFrame(tab_3,text="Show Properties of Avg Log",height=570,width=620)
label_showgraph_tab_3.grid(row=0,column=5,rowspan=1000,columnspan=1000,padx=3,pady=2,sticky=NW)
label_showgraph_tab_3.grid_propagate(False)
label_showgraph_tab_3.config(font=("HP Simplified bold",12),background="white")
###
var_graph_tab_3=StringVar()
var_graph_tab_3.set("pressure")
input_properties_tab_3=OptionMenu(label_showgraph_tab_3,var_graph_tab_3,"pressure","temp ext","fluid velocity","mass rate")
input_properties_tab_3.grid(row=0,column=0,columnspan=1,padx=3,pady=2,sticky=NW)
input_properties_tab_3.config(font=("HP Simplified",9),background="white")
button_excute_showgraph_tab_3=Button(label_showgraph_tab_3,text="Show Graph",command=plot_properties_avg_log)
button_excute_showgraph_tab_3.grid(row=0,column=1,columnspan=1,padx=3,pady=2,sticky=NW)
button_excute_showgraph_tab_3.config(font=("HP Simplified",9),background="white")

## Save GUI
label_save_title=Label(tab_4,text="Save File")
label_save_title.grid(row=0 ,column=0,columnspan=3,padx=3,pady=2,sticky=W)
label_save_title.config(font=("HP Simplified bold",12),background="white")
label_save=Label(tab_4,text="File Name:")
label_save.grid(row=1,column=0,columnspan=1,padx=3,pady=2,sticky=W)
label_save.config(font=("HP Simplified",11),background="white")
input_save=Entry(tab_4,width=16,relief=SOLID)
input_save.grid(row=1,column=1,columnspan=2,padx=3,pady=2,sticky=E)
input_save.config(font=("HP Simplified",10),background="white")
button_excute_save = Button(tab_4,text="Execute",command=save_xlsx)
button_excute_save.grid(row=2,column=2,columnspan=1,padx=3,pady=2,sticky=E)
button_excute_save.config(font=("HP Simplified",9),background="white")

#
root.mainloop()