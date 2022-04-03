from asyncore import read
from cProfile import label
from csv import writer
from multiprocessing.sharedctypes import Value
from operator import index
from statistics import mean
from tkinter import *
import tkinter as tk
from tkinter import ttk, filedialog
from tkinter import messagebox
import tkinter
from tkinter.font import BOLD
from turtle import bgcolor
from unicodedata import numeric
#from matplotlib.font_manager import _Weight
from openpyxl import Workbook
from pyparsing import And
from tkcalendar import DateEntry
import numpy
import pandas as pd
from datetime import datetime
import re
import sys
import os
from tkinter.filedialog import askopenfilename
import matplotlib.pyplot as plt
from datatest import validate
import babel.numbers
import xlsxwriter
import numpy

#import Pmw
#Pmw.initialise()

root = Tk()
root.title("MPI Report")
# screen_width = root.winfo_screenwidth()
# screen_height = root.winfo_screenheight()
root.geometry("965x571")
root.resizable(False,False)
# root.resizable(True,True)
root.configure(bg='#87ceeb')
tabControl = ttk.Notebook(root)
global P_metric
scroll_count = 0
filename = ""

def get_filename():

    global filename, df, lbl1

    
    filename = askopenfilename(filetypes =[('CSV Files', '*.csv'),("All Files","*.*")])    
    if filename == '':
        pass  

    elif not filename.endswith('.csv'):
            messagebox.showinfo("File Type Error", "File must be CSV")  # Alert window
            sys.exit()

    else:
        try:
            df = pd.read_csv(filename) 
        except pd.errors.EmptyDataError:
            messagebox.showinfo("Blank File Notification","File is empty")
            

            #if len(df)<1:
            
            #messagebox.showinfo("Notification","Data not found in file")
           #
        

        if  not set(['DEO Scope','Coding Start Date','Coding Finish Date','Actual Effort']).issubset(df.columns):
            messagebox.showinfo("Notification","required column is not present")
            
        else:
            #messagebox.showinfo("Success", "CSV File imported successfully!!!")
            lbl1.config(text='File imported')  
            btn1['state']="normal"
            btn2['state']="normal"
            clr_data['state']="normal"        

            
        
    
    
    #else:
            # Input = pd.read_csv(filename) 
            # df = Input        
            #messagebox.showinfo("Success", "CSV File imported successfully!!!")
            # lbl.config(text='File imported')

            # btn1['state']="normal"
            # btn2['state']="normal"
            # clr_data['state']="normal"
            #print("get file name",filename)
            #return filename
    # else : 
    #     print(filename)
    #     messagebox.showinfo("File", "Alredy imported")
        
    #return filename


    


    #--------------------------------------------------------------------------------------------------------------------------------


def mpi(): 

    global df,P_metric

    if PM.get()=="":
        messagebox.showinfo("Notification","Planned Metric should not empty")
        
    # elif not(PM.get().isdigit()):
    #         messagebox.showinfo("Notification","Planned Metric should Numeric Inputs"

    elif not (PM.get().replace(".", "", 1).isdigit()):
        messagebox.showinfo("Notification","Planned Metric should Numeric Inputs")      
    
    elif len(PM.get().replace(".", "", 1))>4:
        messagebox.showinfo("Notification","Enter only 4 digits for Planned Mrteic")

    elif entsd.get_date() and ented.get_date()=="":
        messagebox.showinfo("Notification","Start date and End date should be empty")
    
    elif entsd.get_date()>ented.get_date():
        messagebox.showinfo("Notification","Start date should be less than End date")
    
    elif entsd.get_date() == ented.get_date():
        messagebox.showinfo("Notification","Start date and End Date should not same")
    
   
    
    
    
    else:

        global df  

        df["DEO Scope"] = df["DEO Scope"].apply(lambda x: re.sub(r',','', str(x)))
        df["DEO Scope"] = df["DEO Scope"].fillna(0.0).astype('float64')


            #df["DEO Scope"] = df["DEO Scope"].str.replace(',','').astype('float64')
            #df["Actual Effort"] = df["Actual Effort"].str.replace(',','').astype('float64')
            #print(df["DEO Scope"].dtypes)  
                
            #data = pd.DataFrame(df.groupby('Assigned To')['DEO Scope', 'Actual Effort'].agg('sum'))
                


        df['Coding Start Date']= pd.to_datetime(df['Coding Start Date'])
        df['Coding Finish Date']= pd.to_datetime(df['Coding Finish Date'])


        start_date = '{:%Y-%m-%d}'.format(entsd.get_date())
        end_date = '{:%Y-%m-%d}'.format(ented.get_date())
        
        
        mask = (df['Coding Start Date'] >= start_date) & (df['Coding Finish Date'] <= end_date)
        df = df.loc[mask]

              
        #df.rename(columns={'id': 'id_new', 'object': 'object_new'}, inplace=True)
                        
        df.rename(columns={'Assigned To':'Coder Name'}, inplace=True)
        df.rename(columns={'Activity Name':'Activity_Name'}, inplace=True)#SuperLink Counts
        df.rename(columns={'DEO Scope':'Superlink Counts'}, inplace=True)
        df.rename(columns={'Actual Effort':'Coding Hours'}, inplace=True)

                
        #----filterQIL
        df = (df.loc[df.Activity_Name !="DAT QIL"])#exclude by Activity Name


        #----GET planned metric
        P_metric = float(PM.get())
                

        df = pd.DataFrame(df.groupby(['Coder Name'])['Superlink Counts', 'Coding Hours'].agg('sum').reset_index())  

            
        #----MPI/MEtric_Calculation
        df['Metric Achieved'] = df['Superlink Counts']/df['Coding Hours']
        df['MPI'] = df['Metric Achieved']/P_metric
                
            
        #----sorted=ggs.sort_values("Mpi",ascending=False)
                
        required_col = ["Coder Name",'Superlink Counts','Coding Hours','Metric Achieved','MPI']
        df = df[required_col] 


        #----Rounding column values
        df = df.round({'Metric Achieved': 2, 'MPI': 2})


        #---sorting_data
        df=df.sort_values("Metric Achieved",ascending=False)

        
    

        trv(df)    
                # ax=plt.gca()
                # df.plot(kind='bar',x="Name",y="Metric Achieved")
                # plt.show()
                        
                        #messagebox.showinfo("Notification","Data not found in file")

            

            



        # elif 'DEO Scope'in df.columns:
        #     messagebox.showinfo("Notification","")
        # else:
        #     messagebox.showinfo("Notification","Data not found in file")

        


def exportcsv():

    global df

    if len(df) < 1:
        messagebox.showerror("No data", "No data avlaible to export")
        return False

    else:
        writer = filedialog.asksaveasfilename(initialdir=os.getcwd(), title="Save Excel", defaultextension=".xlsx",filetypes=(("Excel File", "*.xlsx"),("ALL Files", "*.*")))
        
        writer = pd.ExcelWriter(writer, engine='xlsxwriter')
        df.to_excel(writer, index=False,sheet_name='Sheet1')

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
        worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(df), len(df.columns)), {'type': 'no_errors', 'format': border_fmt})

        header_fmt=workbook.add_format({'bold':True,'fg_color':'#FFA500'})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num,value, header_fmt)

        writer.save()
        messagebox.showinfo("File Saved", "Your file has been Saved successfully.")
    






#-------------------------------------------------------------------------------------------------------------------------


box1 = tk.LabelFrame(
        root,
        text=" Please Select Date Range ",
        background = "#87ceeb",
        font="TimesNewRoman"
        )


box1.pack(
    ipadx=10,
    ipady=10,
    fill='x'
)
btn=tk.Button(box1,text=("Import File"),font=("TimesNewRoman",10), command=get_filename)
btn.pack(side=tk.LEFT, padx=10, pady=10)

lbl1 = Label(box1,background = "#87ceeb")
lbl1.pack(side=tk.LEFT, padx=10, pady=10,expand=False)


lbl=Label(box1, text = "Start Date:",background = "#87ceeb",font=("TimesNewRoman",10)).pack(side=tk.LEFT, padx=10, pady=10)
entsd=DateEntry(box1,Selectmode="day",width=16)
entsd.delete(0,"end")
entsd.pack(side=tk.LEFT, padx=10, pady=10)


lbled=Label(box1, text = "End Date:",background = "#87ceeb",font=("TimesNewRoman",10)).pack(side=tk.LEFT, padx=10, pady=10)
ented=DateEntry(box1,Selectmode="day",width=16)
ented.delete(0,"end")
ented.pack(side=tk.LEFT, padx=10, pady=10)



        
PM=Label(box1, text = "Enter Planned Metric:",background = "#87ceeb",font=("TimesNewRoman",10)).pack(side=tk.LEFT, padx=10, pady=10)
PM =Entry(box1)
PM.pack(side=tk.LEFT, padx=10, pady=10)


#-------------------------------------------------------------------------------------------------------------------------


box3 = tk.LabelFrame(
        root,
        text=" Select Action to be Taken ",
        background = "#87ceeb",
        font=("TimesNewRoman")
        )


box3.pack(
    ipadx=10,
    ipady=10,
    fill='x'
)


btn1=tk.Button(box3,text="Generate MPI Report",font=("TimesNewRoman",10), command=mpi,state=tk.DISABLED)
btn1.pack(side=tk.LEFT, padx=10, pady=10)

btn2=tk.Button(box3,text="Save MPI Report",font=("TimesNewRoman",10), command=exportcsv,state=tk.DISABLED)
btn2.pack(side=tk.LEFT, padx=10, pady=10)








 
# ----Tab2------------------------------------------------------------------------------------------
tab2=ttk.Frame(tabControl)
tabControl.add(tab2, text ='Report')
tabControl.pack(expand = 1, fill ="both")
ttk.Label(tab2, text ="").pack(expand = 1, fill ="both") 
box4=ttk.LabelFrame(tab2)


#---Box4/Treeview------------------------------------------------------------------------------------

box4.pack(
    ipadx=0,
    ipady=10,
    fill='x',
)


style=ttk.Style()
style.theme_use('clam')
style.configure("W.Treeview", background="#F8F8F8", foreground="black")
style.map("W.Treeview", background=[('selected','green')], foreground=[('selected', 'white')])
#weeklbl=Label(tab2 ,background = "#87ceeb",font=("TimesNewRoman",10))

date_lbl=Label(box4, textvariable=str,background = "white").pack(side=tk.LEFT, padx=450, pady=0)

columns = ("Coder Name","Superlink Counts",'Metric Achieved','MPI')

tree = ttk.Treeview(tab2,columns=columns, show="headings",height=20,style="W.Treeview")
tree.column(1,anchor=CENTER,width=200)
tree['show']='tree'

tree.place(x=0,y=0)
# tree.pack(fill=BOTH)   

def trv(df1):
    
    clear_treeview()
    print(df1)
    counter=0
    global P_metric
    actual_metric=(sum(df1['Superlink Counts'])/sum(df1['Coding Hours'])).__round__(2)
    total_list=['Total']
    total_list.append(sum(df1['Superlink Counts']))
    total_list.append(sum(df1['Coding Hours']))
    total_list.append(actual_metric)
    
    mpi_local= (actual_metric/P_metric).__round__(2)

    total_list.append(mpi_local)

    print(total_list)
    
    tree.tag_configure('g_total',background='silver',font=('TimesNewRoman',10,BOLD))

    # global tree

    #----Add new data in Treeview widget
    tree["column"] = list(df1.columns)
    tree["show"] = "headings"

    #----For Headings iterate over the columns
    for col in tree["column"]:
        tree.heading(col, text=col,anchor= CENTER)

    #----Put Data in Rows
    df_rows = df1.to_numpy().tolist()
    #df_rows = df1.
    print(len(df_rows))

    for row in df_rows:
        tree.insert("", "end", values=row)
        if counter==len(df_rows)-1:
            tree.insert("", "end", values=total_list,tags=['g_total'])
        counter=counter+1
    
   
    tree.column("Coder Name", anchor=CENTER, stretch=YES,width=158)
    tree.column("Superlink Counts", anchor=CENTER, stretch=YES)
    tree.column("Coding Hours", anchor=CENTER, stretch=YES) 
    tree.column("Metric Achieved", anchor=CENTER, stretch=YES) 
    tree.column("MPI", anchor=CENTER, stretch=YES)

    add_scrollbar()
   
def add_scrollbar():
    global scroll_count
    global tree
#print("Before :",scroll_count)
    if scroll_count == 0:
        sb = ttk.Scrollbar(tab2, orient="vertical", command=tree.yview)
        sb.pack(side=RIGHT, expand=YES, fill=BOTH)
        sb1 = ttk.Scrollbar(tab2, orient="horizontal", command=tree.xview)
        sb1.pack(side=BOTTOM, expand=YES, fill=BOTH)
        tree.configure(xscrollcommand=sb1.set)
        tree.configure(yscrollcommand=sb.set)
        tree.pack()
        scroll_count = 1
    else:
        scroll_count = 1
#print("After :",scroll_count)            



#----Clear the Treeview Widget
def clear_treeview():
        
    tree.delete(*tree.get_children())
              

    #----Create a Treeview widget


        # label = Label(box2, text='')
        # label.pack(pady=20)


tab1=ttk.Frame(tabControl)
tabControl.add(tab1, text ='Graph')
tabControl.pack(expand = 1, fill ="both")
ttk.Label(tab1,text ="").pack(expand = 1, fill ="both") 


clr_data=tk.Button(box3,text="Clear Data",font=("TimesNewRoman",10),command=clear_treeview,state=tk.DISABLED)
clr_data.pack(side=tk.LEFT, padx=10, pady=10)


#------------------------------------------------------------------------------------------------------------------------

root.mainloop()
