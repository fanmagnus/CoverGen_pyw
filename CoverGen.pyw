# -*- coding: utf-8 -*-
"""
Generates individualized cover letters in PDF format using settings provided. 
@author: Fan
"""
from configobj import ConfigObj
import tkinter as tk
import tkinter.messagebox as mb
from fpdf import FPDF
import datetime
import os
import webbrowser
class variable(object):
    def __init__(self):
        object.__init__(self)
        self.n=tk.IntVar()
        self.generated=tk.StringVar()
        self.confMain=[]
        self.confMain.append(tk.IntVar())
        for i in range(1,12):
            self.confMain.append(tk.StringVar())
        self.confSec=[]
        for i in range(6):
            self.confSec.append(tk.StringVar())
        self.confName=[]
        for i in range(3):
            self.confName.append(tk.StringVar())
        self.order=[[],[],[]]
        for i in range(3):
            for j in range(10):
                self.order[i].append(tk.IntVar())
        self.format=[]
        for i in range(2):
            self.format.append(tk.IntVar())
        for i in range(2,4):
            self.format.append(tk.StringVar())
        for i in range(4,11):
            self.format.append(tk.BooleanVar())
        
    def readIni(self,i):
        """
        Reads, repairs or creates CLini files, where i is the index of the 
        file. The file CLini0 is the one loads when launched by default.
        """
        existence=1
        corrupt=0
        try:        #Tests if the ith CLini file exists. Loads if it does.
            open("CLini"+str(i))
            config = ConfigObj("CLini"+str(i))
        except:
            existence=0
            config = ConfigObj()
            config.filename = "CLini"+str(i)
        try:
            config["generated"]
        except:
            config["generated"]="firebrick1"
            corrupt=1
        try:
            config["confMain"][11]
        except:
            config['confMain']=[0,"Sample","Recruiters","1. Wall Street","New York, NY, 10001","212.121.2121","","Analyst","Morgan Stanley","indeed.ca","www.ms.com","placeholding text"]
            corrupt=1
        try:
            config["confSec"][5]
        except:
            config['confSec']=["c:\\Users\\User","Simba","Want to contact me?","+1.403.888.3340","fanmagnus@gmail.com",""]
            corrupt=1
        try:
            config["confName"][2]
        except:
            config["confName"]=["Config 1","Config 2","Config 3"]
            corrupt=1
        try:
            config["order0"][9]
        except:
            config["order0"]=[5,4,3,2,1,0,0,0,0,0]
            corrupt=1
        try:
            config["order1"][9]
        except:
            config["order1"]=[0,0,0,0,0,0,0,0,0,0]
            corrupt=1
        try:
            config["order2"][9]
        except:
            config["order2"]=[0,0,0,0,0,0,0,0,0,0]
            corrupt=1
        try:
            config["format"][10]
        except:
            config["format"]=[18,12,"Arial","Arial",1,1,1,0,1,1,1]
            corrupt=1
        for j in range(10):
            try:
                config["para"+str(j)]
            except:
                config["para"+str(j)]="Paragraph " +str(j)+" goes here!"
                corrupt=1                
        """
        For mass letter generation tab. "n" is the number of available entries. 
        "confMainn" is the main configuration of this letter. The first 12 entries
        are the same as "confMain", then label color.
        """
        try:
            config["n"]
        except:
            config["n"]=20
            corrupt=1
        for j in range(int(config["n"])):
            try:
                config["confMainn"+str(j)][12]
            except:
                config["confMainn"+str(j)]=[0,"","","","","","","","","","","","lavender"]
        if existence==0 or corrupt==1:            
            config.write()
            if existence==0:
                mb.showinfo("Setting Created","A new setting has been created.")
            else:
                mb.showinfo("Setting Incomplete","The setting may have been corrupted or outdated. It is now hopefully restored.")
        """
        The data in the CLini file (except above mentioned "n" and "confMainn":
        -"generated" is label color applied for jobs with generated letters.
        -"confname" is the list of paragraph configuration names. e.g., IT or 
        research.
        -"tex" is a list that stores the paragraph texts.
        -"order" contains information of the paragraph orders for all 
        configurations.
        -"confMain" is a list of variables likely to vary from letter to letter: 
        confMain[0] specifies the paragraph order; confMain[1] is the file name;
        confMain[2] is the name of the recipient; confMain[3] through [6] is 
        the recipient's address in 4 lines. confMain[7] through [9] are replacer 
        words in the letter, to replace %x, %y and %z respectively.
        -"confSec" is a list of variables that are more stable:
        confSec[0] specifies the directory for the PDFs. confSec[1] is the
        author's name. confSec[2] through [5] is the author's contact info.
        -"format" stores formating choices: [0] title size, [1] text size, [2] 
        title font, [3] text font, [4] whether to show date in letter, and [5] 
        whether to open file when letter is generated.
        """
        for j in range(11):
            self.format[j].set(config["format"][j])
        for j in range(3):
            self.confName[j].set(config['confName'][j])
        self.para=[]
        for j in range(10):
            self.para.append(config['para'+str(j)])
        for j in range(3):
            for k in range(10):
                self.order[j][k].set(config['order'+str(j)][k])
                self.order[j][k].trace("w",lambda a,b,c,l=self.order[j],index=k:tieBreaker(l,index))
        for j in range(12):
            self.confMain[j].set(config["confMain"][j]) 
        for j in range(6):            
            self.confSec[j].set(config["confSec"][j])
        """
        Reading info for the mass gen letter tab:
        """
        self.generated.set(config["generated"])
        self.n.set(config["n"])
        self.confMainn=[]
        for j in range(self.n.get()):
            self.createNewMainn()
            for k in range(13):
                self.confMainn[j][k].set(config["confMainn"+str(j)][k])
    
    def createNewMainn(self):
        """
        Creates a new entry in self.confMainn:
        """
        self.confMainn.append([tk.IntVar()])
        for k in range(1,13):
            self.confMainn[-1].append(tk.StringVar())
            
class Main(object): 
    def __init__(self):        
        self.interface=tk.Tk()
        self.variable=variable()
        self.currentSave=tk.IntVar()
        self.interface.title("Fan's Cover Letter Generator")
        self.interface.geometry("910x746+100+100")
        self.interface.protocol("WM_DELETE_WINDOW", self.onExit)
        """
        Initiating main interface frames:
        """
        self.interface.menuframe=tk.Frame(self.interface,bg="MediumPurple4")
        self.interface.menuframe.grid(row=0,column=0,sticky="EW",columnspan=2)
        self.interface.tabframe=tk.Frame(self.interface,bg="MediumPurple3")
        self.interface.tabframe.grid(row=2,column=0,sticky="WNS")
        self.interface.manyFrame=tk.Frame(self.interface,bd=10,bg="lavender")
        self.interface.manyFrame.grid(row=2,column=1)
        self.interface.singleFrame=tk.Frame(self.interface,bd=10,bg="lavender")        
        self.interface.textFrame=tk.Frame(self.interface,bd=10,bg="lavender")
        """
        Initiating menuframe:
        """
        numberOfSettings=3 #how many save files are available?
        menus=["Load Settings","Save Settings"] #what buttons appear in the menu bar?
        self.interface.menuButtons=[]
        for text in menus:
            self.interface.menuButtons.append(tk.Menubutton(self.interface.menuframe,text=text,bg="MediumPurple4",activebackground="cornflower blue",relief="raised",bd=5,fg="white",font=("times",13,"bold")))
            self.interface.menuButtons[-1].grid(row=0,column=menus.index(text),sticky="w")
            self.interface.menuButtons[-1].menu=tk.Menu(self.interface.menuButtons[-1],bg="MediumPurple4",activebackground="cornflower blue",fg="white",tearoff="off")
            self.interface.menuButtons[-1]["menu"]=self.interface.menuButtons[-1].menu
        for i in range(numberOfSettings):
            self.interface.menuButtons[0].menu.add_command(label="Setting "+str(i),command=lambda x=i: self.loadIni(x))
            self.interface.menuButtons[1].menu.add_command(label="Setting "+str(i),command=lambda x=i: self.saveIni(x))
        """
        Initiating tabframe:
        """
        self.interface.tabframesub=tk.Frame(self.interface.tabframe,bg="MediumPurple3")
        self.interface.tabframesub.grid(row=0,column=0,pady=80)
        self.interface.buttons=[] #to add a tab button, simply add to the buttonDict
        self.buttonDict={"Main View":self.interface.manyFrame,"Detail Settings":self.interface.singleFrame,"Edit Paragraphs":self.interface.textFrame}
        for text in self.buttonDict:
            self.interface.buttons.append(tk.Button(self.interface.tabframesub,text=text,bg="MediumPurple3",fg="white",activebackground="dodger blue",disabledforeground="black"))
            self.interface.buttons[-1].grid(row=len(self.interface.buttons),column=0,sticky="EW")
            self.interface.buttons[-1].config(command=lambda b=self.interface.buttons[-1],f=self.buttonDict[text]: self.tabSwitch(b,f))
        self.interface.buttons[0]["state"]="disabled"
        self.currentTab=self.interface.buttons[0]                
        """
        Initiating singleFrame:
        """
        self.interface.BGen=tk.Button(self.interface.singleFrame, text ="Generate Letter!", font=('times',20,'bold'),command=lambda:self.genLetter(self.variable.confMain,False), bd=8,bg="purple4",fg="white",activebackground="deep sky blue")
        self.interface.BGen.grid(row=0,column=0,sticky="NEWS",pady=5,columnspan=2) #the button to generate a letter.
        self.interface.singleFrame.about=tk.Message(self.interface.singleFrame,text="Developed by: \nFan Yang, for personal use and free distribution",font=('times',14),width=500,bg="lavender")
        self.interface.singleFrame.about.grid(row=1,column=0,sticky="EW")
        self.interface.singleFrame.add=tk.Button(self.interface.singleFrame,text="Add Recipient Info to Entry List",command=self.addToList,bg="MediumPurple1",fg="white",activebackground="LightBlue1")
        self.interface.singleFrame.add.grid(row=1,column=1,sticky="s")
        """
        Initiating textFrame:
        """
        self.interface.textFrame.instruction=tk.Message(self.interface.textFrame,text="Paragraphs with higher priority print first; only paragraphs with nonzero priority will be printed!",font=('times',12,'bold'),width=500,bg="lavender")
        self.interface.textFrame.instruction.grid(row=0,column=0,columnspan=5,pady=5)
        self.interface.para=[]
        self.interface.textFrame.scrollbar=[]
        self.interface.textFrame.ltx=[]
        for j in range(10):
            self.interface.textFrame.scrollbar.append(tk.Scrollbar(self.interface.textFrame))
            self.interface.textFrame.scrollbar[j].grid(row=j+2,column=5,sticky="NS")
            self.interface.para.append(tk.Text(self.interface.textFrame,height=3,bd=2,width=65,yscrollcommand=self.interface.textFrame.scrollbar[j].set,wrap="word"))
            self.interface.para[j].grid(row=j+2,column=4)
            self.interface.textFrame.scrollbar[j].config(command=self.interface.para[j].yview)            
            self.interface.textFrame.ltx.append(tk.Button(self.interface.textFrame,command=lambda a=j:self.enlarge(a),text="Edit",bg="lavender",activebackground="dodger blue"))
            self.interface.textFrame.ltx[j].grid(row=j+2,column=3)
        self.interface.confName=[] #widgets to select confname
        for j in range (3):
            self.interface.confName.append(tk.Entry(self.interface.textFrame,textvariable=self.variable.confName[j],width=10,bd=2))
            self.interface.confName[j].grid(row=1,column=j)       
        self.interface.order=[[],[],[]] #stores order
        self.temporder=[[],[],[]]
        for j in range(3):
            for k in range(10): 
                self.interface.order[j].append(tk.OptionMenu(self.interface.textFrame,self.variable.order[j][k],0,1,2,3,4,5,6,7,8,9,10))
                self.interface.order[j][k].config(bg="lavender",activebackground="dodger blue")
                self.interface.order[j][k]["menu"].config(bg="lavender",activebackground="dodger blue")
                self.interface.order[j][k].grid(row=k+2,column=j,sticky="EWNS")
        """
        Initiating manyFrame:
        """
        self.interface.BGenMul=tk.Button(self.interface.manyFrame,text ="Generate Letters!",font=('times',20,'bold'),command=self.genMany,bd=8,bg="purple4",fg="white",activebackground="deep sky blue")
        self.interface.BGenMul.grid(row=0,column=0,sticky="NEWS",columnspan=7,pady=5) #the button to generate multiple letters.
        self.interface.bool=tk.BooleanVar()
        self.interface.generated=tk.Menubutton(self.interface.manyFrame,relief="raised",width=25,bd=2)
        self.interface.generated.menu=tk.Menu(self.interface.generated,bg="lavender",tearoff="off")
        self.interface.generated["menu"]=self.interface.generated.menu
        self.interface.colorList=["blue","cyan","gold","firebrick1","black"]
        self.interface.generated.menu.add_command(background="lavender",label="                                        ",command=lambda: self.updateGenColor("lavender"))
        for color in self.interface.colorList:
            self.interface.generated.menu.add_command(background=color,command=lambda: self.updateGenColor(color))
        self.interface.generated.grid(row=1,column=5,pady=5,sticky="w")
        self.interface.manyFrame.AddEntry=tk.Button(self.interface.manyFrame,text="+",bg="lavender",command=self.addLastEntry,fg="green",width=3)
        self.interface.manyFrame.AddEntry.grid(row=1,column=0,sticky="w")
        self.interface.manyFrame.instruction=tk.Label(self.interface.manyFrame,text="Generate letters for checked entries, and set color label to:",justify="right",font=("bold"),bg="lavender")
        self.interface.manyFrame.instruction.grid(row=1,column=1,columnspan=4,sticky="e")
        self.interface.manyFrame.nCount=tk.Label(self.interface.manyFrame,text="No. of Entries: "+str(self.variable.n.get()),bg="lavender")
        self.interface.manyFrame.nCount.grid(row=1,column=6,sticky="e")
        self.interface.manyFrame.manyTitle=[]
        self.interface.manyFrame.manyTitle.append(tk.Checkbutton(self.interface.manyFrame,variable=self.interface.bool,bg="lavender",command=lambda:self.selectAll(self.interface.bool.get())))
        self.interface.manyFrame.manyTitle.append(tk.Button(self.interface.manyFrame,text="Label",width=5,bg="MediumPurple1",fg="white",activebackground="LightBlue1",command=lambda t=12:self.sortList(t)))
        self.interface.manyFrame.manyTitle.append(tk.Button(self.interface.manyFrame,text="%x",width=17,bg="MediumPurple1",fg="white",activebackground="LightBlue1",command=lambda t=7:self.sortList(t)))
        self.interface.manyFrame.manyTitle.append(tk.Button(self.interface.manyFrame,text="File Name",width=17,bg="MediumPurple1",fg="white",activebackground="LightBlue1",command=lambda t=1:self.sortList(t)))
        self.interface.manyFrame.manyTitle.append(tk.Button(self.interface.manyFrame,text="Recipient Name",width=17,bg="MediumPurple1",fg="white",activebackground="LightBlue1",command=lambda t=2:self.sortList(t)))
        self.interface.manyFrame.manyTitle.append(tk.Button(self.interface.manyFrame,text="Notes",width=23,bg="MediumPurple1",fg="white",activebackground="LightBlue1",command=lambda t=11:self.sortList(t)))
        for j in range(6):
            self.interface.manyFrame.manyTitle[j].grid(row=2,column=j,sticky="ew")
        """
        detailframe is the main interface of singleFrame. It consists of detailsub1,
        recframe, formatframe, miscframe and writerframe.
        """
        self.interface.detailframe=tk.Frame(self.interface.singleFrame,bg="lavender")
        self.interface.detailframe.grid(row=2,column=0,pady=30,columnspan=2)
        drawDetail(self.interface.detailframe) #draw detailsub1 and recframe
        """
        Initiating formatframe:
        """
        self.interface.formatframe=tk.LabelFrame(self.interface.detailframe,text="Formating Options",bg="lavender")
        self.interface.formatframe.grid(row=1,column=0,rowspan=2,padx=15)
        self.interface.formatframe.bframe=self.drawFormat(self.interface.formatframe,"Title",2,0,14,24)
        self.interface.formatframe.bframe.grid(row=0,column=0)
        self.interface.formatframe.sframe=self.drawFormat(self.interface.formatframe,"Text",3,1,9,14)
        self.interface.formatframe.sframe.grid(row=1,column=0)
        self.interface.formatframe.topicframe=tk.LabelFrame(self.interface.formatframe,text="Topic Line",bg="lavender")
        self.interface.formatframe.topicframe.grid(row=2,column=0,sticky="ew")
        self.interface.formatframe.topicframe.include=tk.Checkbutton(self.interface.formatframe.topicframe,text="Print topic? If so, include:",variable=self.variable.format[6],bg="lavender")
        self.interface.formatframe.topicframe.include.grid(row=0,column=0,columnspan=4,sticky="w")
        self.interface.formatframe.topicframe.x=tk.Checkbutton(self.interface.formatframe.topicframe,text="%x",variable=self.variable.format[7],bg="lavender")
        self.interface.formatframe.topicframe.x.grid(row=1,column=0,sticky="e",padx=8)
        self.interface.formatframe.topicframe.y=tk.Checkbutton(self.interface.formatframe.topicframe,text="%y",variable=self.variable.format[8],bg="lavender")
        self.interface.formatframe.topicframe.y.grid(row=1,column=1,sticky="e")
        self.interface.formatframe.topicframe.z=tk.Checkbutton(self.interface.formatframe.topicframe,text="%z",variable=self.variable.format[9],bg="lavender")
        self.interface.formatframe.topicframe.z.grid(row=1,column=2,sticky="e",padx=8)
        self.interface.formatframe.topicframe.d=tk.Checkbutton(self.interface.formatframe.topicframe,text="date",variable=self.variable.format[10],bg="lavender")
        self.interface.formatframe.topicframe.d.grid(row=1,column=3,sticky="e")
        self.interface.formatframe.date=tk.Checkbutton(self.interface.formatframe,text="Print date?",variable=self.variable.format[4],bg="lavender")
        self.interface.formatframe.date.grid(row=3,column=0)
        self.interface.formatframe.open=tk.Checkbutton(self.interface.formatframe,text="Open generated letter?",variable=self.variable.format[5],bg="lavender")
        self.interface.formatframe.open.grid(row=4,column=0) 
        """
        Initiating miscframe:
        """
        self.interface.miscframe=tk.LabelFrame(self.interface.detailframe,text="Receiver Details",bg="lavender")
        self.interface.miscframe.grid(row=1,column=1,padx=15)
        self.interface.URLlabel=tk.Label(self.interface.miscframe,text="URL:",anchor="e",bg="lavender")
        self.interface.URLlabel.grid(row=0,column=0,sticky="e")
        self.interface.notelabel=tk.Label(self.interface.miscframe,text="Notes:",anchor="e",bg="lavender")
        self.interface.notelabel.grid(row=1,column=0,sticky="e")
        self.interface.detailframe.ms=[]
        for j in range(2):
            self.interface.detailframe.ms.append(tk.Entry(self.interface.miscframe,textvariable=self.variable.confMain[j+10],bd=2,width=74))
            self.interface.detailframe.ms[j].grid(row=j,column=1,sticky="EW")
        """
        Initiating writerframe:
        """
        self.interface.writerframe=tk.LabelFrame(self.interface.detailframe,text="Writer Information",bg="lavender")
        self.interface.writerframe.grid(row=2,column=1,padx=15,pady=15)
        self.interface.ourframe=tk.LabelFrame(self.interface.writerframe,text="Writer Contact Info",bg="lavender")
        self.interface.ourframe.grid(row=3,column=0,columnspan=2,pady=20)
        self.interface.writerframe.labeldir=tk.Label(self.interface.writerframe,text="Letter Directory:",anchor="e",bg="lavender")
        self.interface.writerframe.labeldir.grid(row=0,column=0)
        self.interface.writerframe.labelname=tk.Label(self.interface.writerframe,text="Enter your name here:",anchor="e",bg="lavender")
        self.interface.writerframe.labelname.grid(row=1,column=0)
        self.interface.confSec=[]
        for j in range(2):
            self.interface.confSec.append(tk.Entry(self.interface.writerframe,textvariable=self.variable.confSec[j],bd=2,width=55))
        self.interface.confSec[0].grid(row=0,column=1,sticky="EW")
        self.interface.confSec[1].grid(row=1,column=1,sticky="EW")
        for j in range(2,6):
            self.interface.confSec.append(tk.Entry(self.interface.ourframe,textvariable=self.variable.confSec[j],width=80,bd=2))
            self.interface.confSec[j].grid(row=j-2,column=0)
        """
        listcanvas is the main interface of the manyFrame.
        """
        self.interface.listscroll=tk.Scrollbar(self.interface.manyFrame)
        self.interface.listcanvas=tk.Canvas(self.interface.manyFrame,bg="lavender",yscrollcommand=self.interface.listscroll.set,width=780,height=550)
        self.interface.listscroll.config(command=self.interface.listcanvas.yview)
        self.interface.listcanvas.grid(row=3,column=0,columnspan=7,sticky="NSEW")
        self.interface.listscroll.grid(row=3,column=7,sticky="NS")
        self.interface.listcanvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.loadIni(0)
        self.interface.generated.config(bg=self.variable.generated.get())
        self.interface.update_idletasks()
        self.interface.listcanvas.config(scrollregion=self.interface.listcanvas.bbox("all"))
        self.updateN()
        self.interface.mainloop()
        
    def _on_mousewheel(self, event):
        self.interface.listcanvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def tabSwitch(self,button,frame):
        """
        Manages tab button activation/deactivations.
        """
        self.currentTab.config(state="normal")
        button.config(state="disabled")
        self.currentTab=button
        switch(frame,2,1) 

    def enlarge(self,i):
        self.interface.secondary=tk.Toplevel(self.interface,bg="lavender")
        self.interface.secondary.title("Edit Paragraph "+str(i))
        self.interface.secondary.geometry("1005x370+100+100")
        self.interface.secondary.id=i
        self.interface.secondary.grab_set()
        self.interface.secondary.scroll=tk.Scrollbar(self.interface.secondary)
        self.interface.secondary.para=tk.Text(self.interface.secondary,height=20,bd=2,width=120,yscrollcommand=self.interface.secondary.scroll.set,wrap="word")
        self.interface.secondary.para.grid(row=0,column=0,sticky="NSEW",pady=5,padx=10)
        self.interface.secondary.scroll.grid(row=0,column=1,sticky="NS")
        self.interface.secondary.scroll.config(command=self.interface.secondary.para.yview)
        self.interface.secondary.para.insert("1.0",self.interface.para[i].get("1.0","end"))
        self.interface.secondary.protocol("WM_DELETE_WINDOW", self.endEnlarge)
        self.interface.secondary.button=tk.Button(self.interface.secondary,text="save",command=self.saveEnlarge)
        self.interface.secondary.button.grid(row=1,column=0)
        
    def saveEnlarge(self):
        self.interface.para[self.interface.secondary.id].delete("1.0","end")
        self.interface.para[self.interface.secondary.id].insert("1.0",self.interface.secondary.para.get("1.0","end"))
  
    def endEnlarge(self):
        if self.interface.secondary.para.get("1.0","end").rstrip()==self.interface.para[self.interface.secondary.id].get("1.0","end").rstrip():
            self.interface.secondary.destroy()
        else:
            t=mb.askyesnocancel("You've made changes!","Would you like to save the changes before leaving?")
            if t:
                self.saveEnlarge()
                self.interface.secondary.destroy()
            elif t is not None:
                self.interface.secondary.destroy()        
        
    def loadIni(self,i):
        self.variable.readIni(i)
        self.currentSave.set(i)
        for j in range(10):
            try:
                self.interface.para[j].delete("1.0","end")
            except:
                pass
            self.interface.para[j].insert("1.0",self.variable.para[j])
        loadDetail(self.interface.detailframe,self.variable.confName,self.variable.confMain)
        try:
            self.listframe.destroy()
        except:
            pass
        
        self.listframe=tk.Frame(self.interface.listcanvas,bg="lavender")
        self.interface.listcanvas.create_window((0,0),window=self.listframe,anchor="nw",width=780)
        self.listframe.items=[[],[],[],[],[],[],[],[],[]]
        self.listframe.checked=[]
        for j in range(self.variable.n.get()):
            self.addEntry(j)

    def saveIni(self,i):
        """
        Saves the current setting to the ith CLini.
        """
        self.currentSave.set(i)
        config = ConfigObj("CLini"+str(i))
        tempconfMain=[]
        tempconfSec=[]
        tempformat=[]
        for j in range(11):
            tempformat.append(self.variable.format[j].get())
        for j in range(6):
            tempconfSec.append(self.variable.confSec[j].get())
        for j in range(12):
            tempconfMain.append(self.variable.confMain[j].get())
        config["confSec"]=tempconfSec.copy()
        config["confMain"]=tempconfMain.copy()
        config["format"]=tempformat.copy()
        para=[]
        temporder0=[]
        temporder1=[]
        temporder2=[]
        for j in range(10):
            para.append(self.interface.para[j].get("1.0","end"))
            config["para"+str(j)]=para[j]
            temporder0.append(self.variable.order[0][j].get())
            temporder1.append(self.variable.order[1][j].get())
            temporder2.append(self.variable.order[2][j].get())
        config["confName"]=[self.variable.confName[0].get(),self.variable.confName[1].get(),self.variable.confName[2].get()]
        config["order0"]=temporder0
        config["order1"]=temporder1
        config["order2"]=temporder2
        config["n"]=self.variable.n.get()
        for j in range(self.variable.n.get()):
            tempMainn=[]
            for k in range(len(self.variable.confMainn[0])):
                tempMainn.append(self.variable.confMainn[j][k].get())
            config["confMainn"+str(j)]=tempMainn.copy()
        config.write() 
        mb.showinfo("Save Successful","Setting "+str(i)+" successfully saved!")
     
    def onLoad(self,i):
        if self.currentSave.get()!=i:
            t=mb.askyesnocancel("Loading New Save","You are about to load a new profile. Would you like to save changes to the current one first?")
            if t:
                self.saveIni(self.currentSave.get())
                self.loadIni(i)
            elif t is not None:
                self.loadIni(i)
        else:
            self.loadIni(i)
            
    def saveExit(self):
        self.saveIni(self.currentSave.get())
        self.interface.destroy()
     
    def onExit(self):
        t=mb.askyesnocancel("You are leaving?","Would you like to save the changes to the current profile?")
        if t:
            self.saveExit()
        elif t is not None:
            self.interface.destroy()
     
    def callThirdly(self,i):
        self.interface.thirdly=tk.Toplevel(self.interface,bg="lavender")
        self.interface.thirdly.title("Details of Entry "+str(i))
        self.interface.thirdly.grab_set()
        self.interface.thirdly.label=tk.Label(self.interface.thirdly,text="URL:",bg="lavender",justify="right")
        self.interface.thirdly.label.grid(row=0,column=0,sticky="e")
        self.interface.thirdly.web=tk.Entry(self.interface.thirdly,textvariable=self.variable.confMainn[i][10],bd=2,width=50)
        self.interface.thirdly.web.grid(row=0,column=1,sticky="w")
        self.interface.thirdly.detailframe=tk.Frame(self.interface.thirdly,bg="lavender")
        self.interface.thirdly.detailframe.grid(row=1,column=0,columnspan=2)
        drawDetail(self.interface.thirdly.detailframe)
        loadDetail(self.interface.thirdly.detailframe,self.variable.confName,self.variable.confMainn[i])     
        
    def genLetter(self,confMain,multiple):
        """
        Generates a letter using given information. multiple is true if generating multiple letters.
        """
        if self.variable.confSec[0].get().rstrip()=="":
            mb.showinfo("File Can't be Created 000","Sorry I can't generate the file. Please provide the directory where you wish the file saved.")
        else:
            Today=datetime.date.today()
            theirAdd=[confMain[3].get(),confMain[4].get(),confMain[5].get(),confMain[6].get()]
            ownAdd=[self.variable.confSec[2].get(),self.variable.confSec[3].get(),self.variable.confSec[4].get(),self.variable.confSec[5].get()]
            para=[]
            for i in range(10):
                para.append(self.interface.para[i].get("1.0","end").rstrip())
            pdf=FPDF()
            pdf.set_author=(self.variable.confSec[1].get())
            pdf.add_page()
            pdf.set_font(self.variable.format[2].get(), "B", size=self.variable.format[0].get())
            pdf.cell(200,10,self.variable.confSec[1].get(),0,1,'C')
            pdf.set_font(self.variable.format[3].get(), "B",size=self.variable.format[1].get())
            ad=''
            for line in ownAdd:
                if len(line)>0:
                    ad=ad+line+' | '
            if len(ad)>3:
                pdf.cell(200,5,ad[0:-3],0,1,'C')
                pdf.ln(8)
            pdf.set_font(self.variable.format[3].get(),size=self.variable.format[1].get())
            if self.variable.format[4].get():
                pdf.write(5,Today.strftime("%B")+" "+str(Today.day)+", "+str(Today.year))
                pdf.ln(8)
            if len(confMain[2].get())>0:
                pdf.write(5,confMain[2].get())
                pdf.ln(5)
            for line in theirAdd:
                if len(line)>0:
                    pdf.write(5,line)
                    pdf.ln(5)
            pdf.ln(5)
            if len(confMain[2].get())>0:
                pdf.write(5,"Dear "+confMain[2].get()+",")
            else:
                pdf.write(5,"Attention: Hiring Manager,")
            if self.variable.format[6].get():                
                re=''
                for j in range(7,10):
                    if self.variable.format[j].get():
                        re=re+confMain[j].get()+', '
                if self.variable.format[10].get():
                    re=re+Today.strftime("%B")+" "+str(Today.day)+", "+str(Today.year)+", "
                if len(re)>0:
                    pdf.ln(8)
                    pdf.set_font(self.variable.format[3].get(), "B",size=self.variable.format[1].get())
                    pdf.write(5,"Re: "+re[0:-2])
                    pdf.set_font(self.variable.format[3].get(),size=self.variable.format[1].get())
            pdf.ln(8)
            seq="" 
            """
            Any paragraph with nonzero priority will appear in the letter, higher 
            priority paraphraphs appear first.
            """
            for i in range(10,0,-1):
                for item in self.variable.order[confMain[0].get()]:
                    if item.get()==i:
                        seq=seq+str(self.variable.order[confMain[0].get()].index(item))
                        break
            for i in seq:
                st=para[int(i)].replace("%x",confMain[7].get())
                st=st.replace("%y",confMain[8].get())
                st=st.replace("%z",confMain[9].get())
                pdf.write(5,st)
                pdf.ln(8)
            pdf.ln(5)
            pdf.write(5,"Sincerely yours,")
            pdf.ln(8)
            pdf.write(5,self.variable.confSec[1].get())
            pdf.ln(10)
            pdf.write(5,"Enc. Resume")
            try:
                pdf.output(self.variable.confSec[0].get()+"\\"+confMain[1].get()+".pdf")
                if self.variable.format[5].get():
                    os.startfile(self.variable.confSec[0].get()+"\\"+confMain[1].get()+".pdf")
            except FileNotFoundError:
                if multiple:
                    raise FileNotFoundError
                else:
                    mb.showinfo("File Can't be Created 001","Sorry I can't generate the file. The file directory provided does not seem to exist.")                    
            except OSError:
                mb.showinfo("File Can't be Created 002","Sorry I can't generate "+confMain[1].get()+".pdf. Maybe the file already exists or there is duplication of file names?")

    def genMany(self):
        """
        Generates letters for all checked entries on manyFrame.
        """
        if self.variable.confSec[0].get().rstrip()=="":
            mb.showinfo("File Can't be Created 000","Sorry I can't generate the files. Please provide the directory where you wish the file saved.")
            
        else:
            color=self.variable.generated.get()
            for i in range(self.variable.n.get()):
                if self.listframe.checked[i].get():
                    try:
                        self.genLetter(self.variable.confMainn[i],True)
                        self.updateLabelColor(i,color)
                    except FileNotFoundError:
                        mb.showinfo("Files Can't be Created 001","Sorry I can't generate the files. The file directory provided does not seem to exist.")
                        break
    
    def sortList(self,i):
        """
        Sort entries on manyFrame.
        """
        self.selectAll(0)
        self.variable.confMainn.sort(key=lambda x:(x[i].get()=="",x[i].get()))
        for j in range(self.variable.n.get()):
            self.listframe.items[2][j].config(textvariable=self.variable.confMainn[j][7])
            self.listframe.items[3][j].config(textvariable=self.variable.confMainn[j][1])
            self.listframe.items[4][j].config(textvariable=self.variable.confMainn[j][2])
            self.listframe.items[5][j].config(textvariable=self.variable.confMainn[j][11])
            self.listframe.items[7][j].config(command=lambda x=self.variable.confMainn[j]:webbrowser.open(x[10].get(),new=2))
            self.updateLabelColor(j,self.variable.confMainn[j][12].get())
    
    def addEntry(self,i):
        """
        Add the ith entry to manyFrame:
        """
        self.listframe.checked.append(tk.BooleanVar())
        self.listframe.items[0].append(tk.Checkbutton(self.listframe,variable=self.listframe.checked[i],width=1,bg="lavender"))
        self.listframe.items[1].append(tk.Menubutton(self.listframe,bg=self.variable.confMainn[i][12].get(),width=7,relief="raised"))
        self.listframe.items[1][i].menu=tk.Menu(self.listframe.items[1][i],bg="lavender",activebackground="cornflower blue",fg="white",tearoff="off")
        self.listframe.items[1][i]["menu"]=self.listframe.items[1][i].menu
        self.listframe.items[1][i].menu.add_command(background="lavender",label="   ",command=lambda x=i,y="lavender": self.updateLabelColor(x,y))
        for color in self.interface.colorList:
            self.listframe.items[1][i].menu.add_command(background=color,command=lambda x=i,y=color: self.updateLabelColor(x,y))
        self.listframe.items[2].append(tk.Entry(self.listframe,textvariable=self.variable.confMainn[i][7],bd=2,width=22))
        self.listframe.items[3].append(tk.Entry(self.listframe,textvariable=self.variable.confMainn[i][1],bd=2,width=22))
        self.listframe.items[4].append(tk.Entry(self.listframe,textvariable=self.variable.confMainn[i][2],bd=2,width=22))
        self.listframe.items[5].append(tk.Entry(self.listframe,textvariable=self.variable.confMainn[i][11],bd=2,width=29))
        self.listframe.items[6].append(tk.Button(self.listframe,text="Edit",bg="lavender",activebackground="LightBlue1",command=lambda x=i: self.callThirdly(x),width=4))
        self.listframe.items[7].append(tk.Button(self.listframe,text="Web",bg="lavender",activebackground="LightBlue1",command=lambda x=self.variable.confMainn[i]: webbrowser.open(x[10].get(),new=2)))
        self.listframe.items[8].append(tk.Button(self.listframe,text="-",fg="red",bg="lavender",activebackground="LightBlue1",command=lambda x=i:self.dropEntry(x),width=2))
        for j in range(9):
            self.listframe.items[j][i].grid(row=i,column=j)
        self.interface.update_idletasks()
        self.interface.listcanvas.config(scrollregion=self.interface.listcanvas.bbox("all"))
     
    def addLastEntry(self):
        """
        Creates an empty entry and add to the end of manyFrame:
        """
        n=self.variable.n.get()
        self.variable.n.set(n+1)
        self.variable.createNewMainn()
        self.variable.confMainn[-1][12].set("lavender")
        self.addEntry(n)
        self.updateN()
        
    def dropEntry(self,i):
        n=self.variable.n.get()
        for j in range(i,n-1):
            self.listframe.checked[j].set(self.listframe.checked[j+1].get())
            for k in range(len(self.variable.confMainn[0])):
                self.variable.confMainn[j][k].set(self.variable.confMainn[j+1][k].get())
            self.updateLabelColor(j,self.variable.confMainn[j][12].get())
        for j in range(len(self.listframe.items)):
            self.listframe.items[j][-1].destroy()
            self.listframe.items[j].pop()
        self.interface.update_idletasks()
        self.interface.listcanvas.config(scrollregion=self.interface.listcanvas.bbox("all"))
        self.variable.confMainn.pop()
        self.updateN()
    
    def addToList(self):
        """
        Creates a new entry then adds self.confMain to it:
        """
        self.addLastEntry()
        for i in range(12):
            self.variable.confMainn[-1][i].set(self.variable.confMain[i].get())
    
    def updateN(self):
        n=len(self.variable.confMainn)
        self.variable.n.set(n)
        self.interface.manyFrame.nCount.config(text="No. of Entries: "+str(n))
        
    def updateLabelColor(self,i,j):
        self.variable.confMainn[i][12].set(j)
        self.listframe.items[1][i].config(bg=j)
        
    def updateGenColor(self,j):
        self.variable.generated.set(j)
        self.interface.generated.config(bg=j) 
        
    def selectAll(self,bool):
        for i in range(len(self.listframe.checked)):
            self.listframe.checked[i].set(bool)
    
    def drawFormat(self,frame,name,fontVariable,sizeVariable,sizeMin,sizeMax):
        labelFrame=tk.LabelFrame(frame,text=name+" Formating",bg="lavender")
        labelFrame.labelSize=tk.Label(labelFrame,text=name+" Font Size:",anchor="e",bg="lavender")
        labelFrame.labelSize.grid(row=0,column=0,pady=5)
        labelFrame.labelFont=tk.Label(labelFrame,text=name+" Font:",anchor="e",bg="lavender")
        labelFrame.labelFont.grid(row=1,column=0,pady=5)
        labelFrame.font=tk.OptionMenu(labelFrame,self.variable.format[fontVariable],"Courier","Helvetica","Arial","Times")
        labelFrame.font.config(bg="lavender",activebackground="dodger blue")
        labelFrame.font["menu"].config(bg="lavender",activebackground="dodger blue")
        labelFrame.font.grid(row=1,column=1,sticky="EW")
        labelFrame.size=tk.Spinbox(labelFrame,textvariable=self.variable.format[sizeVariable],from_=sizeMin,to=sizeMax,bd=2)
        labelFrame.size.grid(row=0,column=1)
        return labelFrame
    
def tieBreaker(list, index):
    """
    This function is used to define main.variable.order. This function 
    makes sure that whenever a paragraph is given a nonzero priority, any 
    paragraph with the same priority is set to zero priority.
    """
    if list[index].get()!=0:
        for i in range(len(list)):
            if list[i].get()==list[index].get() and i!=index:
                list[i].set(0)
                break

def switch(frame,row,column):
    """
    Function that changes tabs in the interface. frame is the tab to be switched
    on, row and column is the location of said tab in its master frame.
    """
    frame.grid(row=row,column=column,sticky="EWNS")
    frame.lift()
    
def drawDetail(frame):
    """
    Initiating detailsub1:
    """
    frame.detailsub1=tk.Frame(frame,bg="lavender")
    frame.detailsub1.grid(row=0,column=0)
    frame.labelframe=tk.LabelFrame(frame.detailsub1,text="Replacable Phrases",bg="lavender")
    frame.labelframe.grid(row=0,column=0)
    frame.labelx=tk.Label(frame.labelframe,text="=%x",bg="lavender")
    frame.labelx.grid(row=0,column=1)
    frame.labely=tk.Label(frame.labelframe,text="=%y",bg="lavender")
    frame.labely.grid(row=1,column=1)
    frame.labelz=tk.Label(frame.labelframe,text="=%z",bg="lavender")
    frame.labelz.grid(row=2,column=1)
    frame.labelframe1=tk.LabelFrame(frame.detailsub1,text="Paragraph Order",bg="lavender")
    frame.labelframe1.grid(row=1,column=0)
    """
    Initiating recframe:
    """
    frame.recframe=tk.LabelFrame(frame,text="Recipient Information",bg="lavender")
    frame.recframe.grid(row=0,column=1)        
    frame.labelfile=tk.Label(frame.recframe,text="Save letter as:",anchor="e",bg="lavender")
    frame.labelfile.grid(row=0,column=0,sticky="e")
    frame.labelfile1=tk.Label(frame.recframe,text=".pdf",anchor="w",bg="lavender")
    frame.labelfile1.grid(row=0,column=2,sticky="w")
    frame.labelrec=tk.Label(frame.recframe,text="Name of recipient:",anchor="e",bg="lavender")
    frame.labelrec.grid(row=1,column=0,sticky="e")
    frame.theirframe=tk.LabelFrame(frame.recframe,text="Recipient Address",bg="lavender")
    frame.theirframe.grid(row=3,column=0,columnspan=3,pady=20)
    frame.rp=[] #widgets for the replacer
    for j in range(3):
        frame.rp.append(tk.Entry(frame.labelframe,bd=2))
        frame.rp[j].grid(row=j, column=0)
    frame.cMain=[] #widget for confMain
    for j in range(2):
        frame.cMain.append(tk.Entry(frame.recframe,bd=2))
    frame.cMain[0].grid(row=0,column=1,sticky="EW")
    frame.cMain[1].grid(row=1,column=1,columnspan=2,sticky="EW")
    for j in range(2,6):
        frame.cMain.append(tk.Entry(frame.theirframe,width=80,bd=2))
        frame.cMain[j].grid(row=j-2,column=0)
    frame.cr=[] #widgets to select con 
    for j in range(3):
        frame.cr.append(tk.Radiobutton(frame.labelframe1,value=j,bg="lavender"))
        frame.cr[j].grid(row=j,column=0,sticky="W")
            
def loadDetail(frame,confName,confMain):
    for i in range(3):
        frame.rp[i].config(textvariable=confMain[i+7])
        frame.cr[i].config(textvariable=confName[i],variable=confMain[0])
    for i in range(6):
        frame.cMain[i].config(textvariable=confMain[i+1])
    for i in range(2):
        frame.ms[i].config(textvariable=confMain[i+10])
                
CL=Main()