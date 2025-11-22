from tkinter import *
from tkinter import ttk
import math
import numpy as np
import pandas as pd
from tkinter import messagebox
import openpyxl, xlrd
from openpyxl import Workbook
import pathlib 



root = Tk() 
root.configure(bg='#262626')
root.geometry("700x580")
root.resizable(False,False)
root.title("Jadwal Shalat 1.0")
gambar = PhotoImage(file = "C:/Users/Tafa/OneDrive/Dokumen/Jadwal Shalat 1.0/bg.png")
label = ttk.Label(root, image=gambar)
PhotoImage(file = "C:/Users/Tafa/OneDrive/Dokumen/Jadwal Shalat 1.0/bg.png")
label.grid(row=1,column=0) 

def toggle_win():
    global f1
    f1=Frame(root,width=250,height=580,bg='#12c4c0')
    f1.place(x=0,y=0)
    
   
    def bttn(x,y,text,bcolor,fcolor,cmd):
     
        def on_entera(e):
            myButton1['background'] = bcolor 
            myButton1['foreground']= '#262626'  
        def on_leavea(e):
            myButton1['background'] = fcolor
            myButton1['foreground']= '#262626'
        myButton1 = Button(f1,text=text,
                       width=42,
                       height=2,
                       fg='#262626',
                       border=0,
                       bg=fcolor,
                       activeforeground='#262626',
                       activebackground=bcolor,            
                        command=cmd)
                      
        myButton1.bind("<Enter>", on_entera)
        myButton1.bind("<Leave>", on_leavea)
        myButton1.place(x=x,y=y)
    bttn(0,80,'H O M E','#0f9d9a','#12c4c0',home)
    bttn(0,117,'H E L P','#0f9d9a','#12c4c0',bantuan)
    bttn(0,154,'A B O U T','#0f9d9a','#12c4c0', definisi)
    

    def dele():
        f1.destroy()
        b2=Button(root,image=img1,
               command=toggle_win,
               border=0,
               bg='#262626',
               activebackground='#262626')
        b2.place(x=5,y=8)
    global img2
    img2 = PhotoImage(file = "C:/Users/Tafa/OneDrive/Dokumen/Jadwal Shalat 1.0/tutup.png")
    Button(f1,
           image=img2,
           border=0,
           command=dele,
           bg='#12c4c0',
           activebackground='#12c4c0').place(x=5,y=10)
img1 = PhotoImage(file = "C:/Users/Tafa/OneDrive/Dokumen/Jadwal Shalat 1.0/buka.png")
global b2
b2=Button(root,image=img1,
       command=toggle_win,
       border=0,
       bg='#262626',
       activebackground='#262626')
b2.place(x=5,y=8)

def default_home():
    l=Label(root,text='Program Waktu Shalat',fg='white',bg='#262626')
    l.config(font=('Comic Sans MS',40))
    l.place(x=75,y=320)
   
default_home() 

def home():
    f1.destroy()
    f2=Frame(root,width=700,height=550,bg='#262626')
    f2.place(x=0,y=45)
    l=Label(f2,text='Jadwal Shalat Selama 1 bulan',fg='white',bg='#262626')
    l.config(font=('Comic Sans MS',10))
    l.place(x=250,y=140)
    cb_label = ttk.Label(root,text="Lokasi Kota")
    cb_label.place(x=85, y = 65)

    pilihan = ["Bogor","Denpasar","Jakarta","Malang",
           "Manado","Makasar","Medan","Mojokerto", 
           "Natuna", "Gorontalo"]
    cb = ttk.Combobox(root,values=pilihan, width=29,height=5)
    cb.set("Pilih Lokasi Kota")
    cb.place(x=30, y=85 )
    
    tanggal_label = ttk.Label(root,text="Tanggal :")
    tanggal_label.place(x = 30, y =115)
    spinbox1 = Spinbox(root,from_=1, to=31)
    spinbox1.place(x = 90, y =115)
    
    bulan_label = ttk.Label(root,text="Bulan     :")
    bulan_label.place(x=30,y =137)
    spinbox2 = Spinbox(root,values=("Pilih Bulan"))
    spinbox2.place(x = 90, y =137)

    tahun_label = ttk.Label(root,text="Tahun    :")
    tahun_label.place(x = 30, y =159)
    spinbox3 = Spinbox(root,from_=1000, to=1000000)
    spinbox3.place(x = 90, y =159)
    # bagian cari masehi-hijriyah
    tombolx = ttk.Button(root,text="Download")
    tombolx.place(x = 330, y =150)
    tombolx = ttk.Button(root,text="Cari")
    tombolx.place(x = 330, y =120)
    


    def a():
      if (c1.get()==1) & (c2.get()==0) :
        spinbox2 = Spinbox(root,values=("Januari","Februari","Maret","April",
                                 "Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"))
        spinbox2.place(x = 90, y =137)
      elif (c1.get()==0) & (c2.get()==1) :
        spinbox2 = Spinbox(root,values=("Muharram","Shafar","Rabi'ul Awal","Rabi'ul akhir",
                                 "Jumadil Awal","Jumadil Akhir","Rajab","Sya'ban","Ramadhan","Syawwal","Zulqaidah","Zulhijjah"))
        spinbox2.place(x = 90, y =137)
      #Database
      csv_data = pd.read_csv("C:/Users/Tafa/OneDrive/Dokumen/waktu sholat/Data_kota.csv", sep=';')
      def tekan() :
        lokasikota = cb.get()
        D = int(spinbox1.get())
        M = spinbox2.get()
        M1 = spinbox2.get()
        Y = int (spinbox3.get())
        if (cb.get() == "Jakarta") :
          L = csv_data['Lintang'].iloc[0]
          BU = csv_data['Bujur'].iloc[0]
          Z = csv_data['Zona Waktu'].iloc[0]
          H = csv_data['Ketinggian'].iloc[0]
        elif (cb.get() ==  "Mojokerto") :
          L = csv_data['Lintang'].iloc[1]
          BU = csv_data['Bujur'].iloc[1]
          Z = csv_data['Zona Waktu'].iloc[1]
          H = csv_data['Ketinggian'].iloc[1]
        elif (cb.get() == "Manado"):
          L = csv_data['Lintang'].iloc[2]
          BU = csv_data['Bujur'].iloc[2]
          Z = csv_data['Zona Waktu'].iloc[2]
          H = csv_data['Ketinggian'].iloc[2]
        elif (cb.get() ==  "Medan"):
          L = csv_data['Lintang'].iloc[3]
          BU = csv_data['Bujur'].iloc[3]
          Z = csv_data['Zona Waktu'].iloc[3]
          H = csv_data['Ketinggian'].iloc[3]
        elif (cb.get() ==  "Natuna") :
          L = csv_data['Lintang'].iloc[4]
          BU = csv_data['Bujur'].iloc[4]
          Z = csv_data['Zona Waktu'].iloc[4]
          H = csv_data['Ketinggian'].iloc[4]
        elif (cb.get() == "Malang") :
          L = csv_data['Lintang'].iloc[5]
          BU = csv_data['Bujur'].iloc[5]
          Z = csv_data['Zona Waktu'].iloc[5]
          H = csv_data['Ketinggian'].iloc[5]
        elif (cb.get() ==  "Makasar") :
          L = csv_data['Lintang'].iloc[6]
          BU = csv_data['Bujur'].iloc[6]
          Z = csv_data['Zona Waktu'].iloc[6]
          H = csv_data['Ketinggian'].iloc[6]
        elif (cb.get() ==  "Bogor") :
          L = csv_data['Lintang'].iloc[7]
          BU = csv_data['Bujur'].iloc[7]
          Z = csv_data['Zona Waktu'].iloc[7]
          H = csv_data['Ketinggian'].iloc[7]
        elif (cb.get() ==  "Denpasar") :
          L = csv_data['Lintang'].iloc[8]
          BU = csv_data['Bujur'].iloc[8]
          Z = csv_data['Zona Waktu'].iloc[8]
          H = csv_data['Ketinggian'].iloc[8]
        elif (cb.get() ==  "Gorontalo") :
          L = csv_data['Lintang'].iloc[9]
          BU = csv_data['Bujur'].iloc[9]
          Z = csv_data['Zona Waktu'].iloc[9]
          H = csv_data['Ketinggian'].iloc[9] 
      
        if (M =="Januari") :
         M = 1
         n = 31
        elif (M =="Februari") :
          M = 2
          if Y%4 == 0 :
            n = 29
          else :
            n = 28
        elif (M =="Maret") :
          M = 3
          n =31
        elif (M =="April") :
          M = 4
          n = 30
        elif (M =="Mei") :
          M = 5
          n = 31
        elif (M=="Juni") :
          M = 6
          n =30
        elif (M=="Juli") :
          M = 7
          n = 31
        elif (M=="Agustus") :
          M = 8
          n =31
        elif (M=="September") :
          M = 9
          n = 30
        elif (M=="Oktober") :
          M = 10
          n = 31
        elif (M=="November") :
          M = 11
          n = 30
        elif (M =="Desember") :
          M = 12
          n = 31

        if (M=="Muharram") :
          M = M1
          M1 = 1
          n = 30
        elif (M1=="Shafar") :
          M = M
          M1 = 2
          n = 29
        elif (M1=="Rabi'ul Awal") :
          M = M
          M1 = 3
          n = 30
        elif (M1=="Rabi'ul akhir") :
          M = M
          M1 = 4
          n = 29
        elif (M1=="Jumadil Awal") :
          M = M1
          M1 = 5
          n = 30
        elif (M1=="Jumadil Akhir") :
          M = M1
          M1 = 6
          n =29
        elif (M1=="Rajab") :
          M = M1
          M1 = 7
          n = 30
        elif (M1=="Sya'ban") :
          M = M1
          M1 = 8
          n = 29
        elif (M1=="Ramadhan") :
          M = M1
          M1= 9
          n = 30
        elif (M1=="Syawwal") :
          M = M1
          M1 = 10
          n = 29
        elif (M1=="Zulqaidah") :
          M = M1
          M1 = 11
          n = 30
        elif (M1=="Zulhijjah") :
          M = M1
          M1 = 12
          if Y%30 == 2 :
            n = 30
          elif Y%30 == 5 :
            n = 30
          elif Y%30 == 7 :
            n = 30
          elif Y%30 == 10 :
            n = 30
          elif Y%30 == 13 :
            n = 30
          elif Y%30 == 16 :
            n = 30
          elif Y%30 == 18 :
            n = 30
          elif Y%30 == 21 :
            n = 30
          elif Y%30 == 24 :
            n = 30
          elif Y%30 == 26 :
            n = 30
          elif Y%30 == 29 :
            n = 30
          else :
            n =29 

        if (c1.get()==1) & (c2.get()==0) :
          
          if M > 2 :
            M = M
            Y = Y
          elif M in [1 or 2]:
            M = M +12
            Y = Y -1
          #masehi
          A  = int(Y/100)
          B = 2 + int(A/4)-A
          JD = 1720994.5 + int(365.25*Y) + int(30.6001*(M+1))+ B + D +0.5
          JD_lokal = JD - Z/24
          T = 2*3.14159265359*(JD_lokal-2451545)/365.25
          Delta = 0.37877 + 23.264*math.sin(math.radians(57.297*T-79.547)) + 0.3812*math.sin(math.radians(2*57.297*T-82.682)) + 0.17132*math.sin(math.radians (3*57.297*T-59.722))
          U = (JD_lokal - 2451545)/36525
          L0x = 280.46607 + 36000.7698*U
          L0 = L0x%360
          rad = np.pi/180
          et = ((-(1789+237*U))*(np.sin(L0*rad)))-((7146-62*U)*(np.cos(L0*rad)))+((9934-14*U)*(np.sin(2*L0*rad)))-((29+5*U)*(np.cos(2*L0*rad)))+((74+10*U)*(np.sin(3*L0*rad)))+((320-4*U)*(np.cos(3*L0*rad)))-((212*np.sin(4*L0*rad)))
          ET = et/1000
          Transit = 12 + Z -(BU/15) - (ET/60)
         # Dzuhur
          zhuhur = Transit # +2menit
          h = (int(zhuhur)%24)
          m = ((zhuhur*60)%60)+2
          s = (zhuhur*3600)%60
          Dzuhur = "%d:%02d:%02d"%(h,m,s)
      #ASHAR
          KA = 1 # syafii
          h_AS=(math.atan(1/(KA+math.tan((abs(Delta-L))*(math.pi/180)))))*180/math.pi
          ASHAR = (math.sin(math.radians(h_AS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
          HA_AS = math.degrees(math.acos(ASHAR))
          "waktu ashar"
          ashar = Transit + (HA_AS)/15
          h = (int(ashar)%24)
          m = ((ashar*60)%60)
          s = (ashar*3600)%60
          Ashar = "%d:%02d:%02d"%(h,m,s)
      #Magrib
          h_MG = -0.8333-00.0347*math.sqrt(H)
          MAGRIB = (math.sin(math.radians(h_MG))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
          HA_MG = math.degrees(math.acos(MAGRIB))
          "waktu magrib"
          magrib = Transit + (HA_MG)/15
          h = (int(magrib)%24)
          m = ((magrib*60)%60)
          s = (magrib*3600)%60
          Magrib = "%d:%02d:%02d"%(h,m,s)
      #Isya
          h_IS = -18
          ISYA = (math.sin(math.radians(h_IS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
          HA_IS = math.degrees(math.acos(ISYA))
          "waktu isya"
          isya = Transit + (HA_IS)/15
          h = (int(isya)%24)
          m = ((isya*60)%60)
          s = (isya*3600)%60
          Isya = "%d:%02d:%02d"%(h,m,s)
      #Shubuh
          h_SH = -20
          SHUBUH = (math.sin(math.radians(h_SH))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
          HA_SH = math.degrees(math.acos(SHUBUH))
          "waktu shubuh"
          shubuh = Transit - (HA_SH)/15
          h = (int(shubuh)%24)
          m = ((shubuh*60)%60)
          s = (shubuh*3600)%60
          Shubuh = "%d:%02d:%02d"%(h,m,s)
      #Terbit
          "waktu terbit"
          terbit = Transit - (HA_MG)/15
          h = (int(terbit)%24)
          m = ((terbit*60)%60)
          s = (terbit*3600)%60
          Wter = "%d:%02d:%02d"%(h,m,s)
          jadwal1 = f"Jadwal Shalat \n {lokasikota}, {spinbox1.get()} {spinbox2.get()} {spinbox3.get()}"  f"\nShubuh : {Shubuh}"  f"\nTerbit : {Wter}" f"\nDzuhur : {Dzuhur}" f"\nAshar : {Ashar}" f"\nMagrib : {Magrib}" f"\nIsya : {Isya}"
          teks.configure(text=jadwal1)
          for i in range(0,n):
            i = i+1
            JD = 1720994.5 + int(365.25*Y) + int(30.6001*(M+1))+ B + i +0.5
            JD_lokal = JD - Z/24
            T = 2*3.14159265359*(JD_lokal-2451545)/365.25
            Delta = 0.37877 + 23.264*math.sin(math.radians(57.297*T-79.547)) + 0.3812*math.sin(math.radians(2*57.297*T-82.682)) + 0.17132*math.sin(math.radians (3*57.297*T-59.722))
            U = (JD_lokal - 2451545)/36525
            L0x = 280.46607 + 36000.7698*U
            L0 = L0x%360
            rad = np.pi/180
            et = ((-(1789+237*U))*(np.sin(L0*rad)))-((7146-62*U)*(np.cos(L0*rad)))+((9934-14*U)*(np.sin(2*L0*rad)))-((29+5*U)*(np.cos(2*L0*rad)))+((74+10*U)*(np.sin(3*L0*rad)))+((320-4*U)*(np.cos(3*L0*rad)))-((212*np.sin(4*L0*rad)))
            ET = et/1000
            Transit = 12 + Z -(BU/15) - (ET/60)
            # Dzuhur
            zhuhur = Transit # +2menit
            h = (int(zhuhur)%24)
            m = ((zhuhur*60)%60)+2
            s = (zhuhur*3600)%60
            Dzuhur = "%d:%02d:%02d"%(h,m,s)
            #ASHAR
            KA = 1 # syafii
            h_AS=(math.atan(1/(KA+math.tan((abs(Delta-L))*(math.pi/180)))))*180/math.pi
            ASHAR = (math.sin(math.radians(h_AS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_AS = math.degrees(math.acos(ASHAR))
            "waktu ashar"
            ashar = Transit + (HA_AS)/15
            h = (int(ashar)%24)
            m = ((ashar*60)%60)
            s = (ashar*3600)%60
            Ashar = "%d:%02d:%02d"%(h,m,s)
            #Magrib
            h_MG = -0.8333-00.0347*math.sqrt(H)
            MAGRIB = (math.sin(math.radians(h_MG))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_MG = math.degrees(math.acos(MAGRIB))
            "waktu magrib"
            magrib = Transit + (HA_MG)/15
            h = (int(magrib)%24)
            m = ((magrib*60)%60)
            s = (magrib*3600)%60
            Magrib = "%d:%02d:%02d"%(h,m,s)
            #Isya
            h_IS = -18
            ISYA = (math.sin(math.radians(h_IS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_IS = math.degrees(math.acos(ISYA))
            "waktu isya"
            isya = Transit + (HA_IS)/15
            h = (int(isya)%24)
            m = ((isya*60)%60)
            s = (isya*3600)%60
            Isya = "%d:%02d:%02d"%(h,m,s)
            #Shubuh
            h_SH = -20
            SHUBUH = (math.sin(math.radians(h_SH))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_SH = math.degrees(math.acos(SHUBUH))
            "waktu shubuh"
            shubuh = Transit - (HA_SH)/15
            h = (int(shubuh)%24)
            m = ((shubuh*60)%60)
            s = (shubuh*3600)%60
            Shubuh = "%d:%02d:%02d"%(h,m,s)
            #Terbit
            "waktu terbit"
            terbit = Transit - (HA_MG)/15
            h = (int(terbit)%24)
            m = ((terbit*60)%60)
            s = (terbit*3600)%60
            Wter = "%d:%02d:%02d"%(h,m,s)
            data = f"{i} {spinbox2.get()}"
            my_tree.insert(parent='',index='end', text= data, values= (f"{Shubuh}",f"{Wter}" ,f"{Dzuhur}" ,f"{Ashar}" ,f"{Magrib}" ,f"{Isya}"))
        elif  (c1.get()==0) & (c2.get()==1) :
          leap_year = [2,5,7,10,13,16,18,21,24,26,29]
          year2 = (Y-1)%30
          if year2 == 2:
            length = len(leap_year[0:1])
          elif 5 <= year2 <7 :
            length = len(leap_year[0:2])
          elif 7 <= year2 < 10 :
            length = len(leap_year[0:3])
          elif 10 <= year2 < 13 :
            length = len(leap_year[0:4])
          elif 13 <= year2 < 16:
            length = len(leap_year[0:5])
          elif 16<= year2 < 18 :
            length = len(leap_year[0:6])
          elif 18<= year2 < 21 :
            length = len(leap_year[0:7])
          elif 21<= year2 < 24 :
            length = len(leap_year[0:8])
          elif 24 <= year2 < 26 :
            length = len(leap_year[0:9])
          elif 26 <= year2 < 29:
            length = len(leap_year[0:10])
          elif 29 <= year2 >= 30 :
            length = len(leap_year[0:11])
          else : 
            length = 0
          year = int((Y - 1)/30)
          total_day = (year*(354*30+11)) + (year2*354 + length)
          hijriah = { 1: 30, 2 : 29, 3 : 30, 4 : 29, 5: 30, 6 : 29, 7 : 30, 8 : 29, 9 : 30, 10 : 29, 11 : 30, 12 : 29}
          key = list(hijriah.keys())
          val = list(hijriah.values())
      
          if M1 == 2:
            sum = val[0]
          elif M1 == 3 :
            sum = val[0]+val[1]
          elif M1 == 4 :
            sum = val[0]+val[1]+val[2]
          elif M1 == 5 :
            sum = val[0]+val[1]+val[2]+val[3]
          elif M1 == 6 :
            sum = val[0]+val[1]+val[2]+val[3]+val[4]
          elif M1 == 7 :
            sum = val[0]+val[1]+val[2]+val[3]+val[4]+val[5]
          elif M1 == 8 :
            sum = val[0]+val[1]+val[2]+val[3]+val[4]+val[5]+val[6]
          elif M1 == 9 :
            sum = val[0]+val[1]+val[2]+val[3]+val[4]+val[5]+val[6]+val[7]
          elif M1 == 10 : 
            sum = val[0]+val[1]+val[1]+val[2]+val[3]+val[4]+val[8]+val[9]+val[10]
          elif M1 == 11 :
            sum = val[0]+val[1]+val[1]+val[2]+val[3]+val[4]+val[8]+val[9]+val[10]+val[11]
          elif M1 == 12 : 
            sum = val[0]+val[1]+val[1]+val[2]+val[3]+val[4]+val[8]+val[9]+val[10]+val[11]+val[11]
          else :
            sum = 0
          month = sum + D
          total = total_day + month
          jd = 1948438.5 + total
          JD_lokal = jd - Z/24
          T = 2*3.14159265359*(JD_lokal-2451545)/365.25
          Delta = 0.37877 + 23.264*math.sin(math.radians(57.297*T-79.547)) + 0.3812*math.sin(math.radians(2*57.297*T-82.682)) + 0.17132*math.sin(math.radians (3*57.297*T-59.722))
          U = (JD_lokal - 2451545)/36525
          L0x = 280.46607 + 36000.7698*U
          L0 = L0x%360
          rad = np.pi/180
          et = ((-(1789+237*U))*(np.sin(L0*rad)))-((7146-62*U)*(np.cos(L0*rad)))+((9934-14*U)*(np.sin(2*L0*rad)))-((29+5*U)*(np.cos(2*L0*rad)))+((74+10*U)*(np.sin(3*L0*rad)))+((320-4*U)*(np.cos(3*L0*rad)))-((212*np.sin(4*L0*rad)))
          ET = et/1000
          Transit = 12 + Z -(BU/15) - (ET/60)
         # Dzuhur
          zhuhur = Transit # +2menit
          h = (int(zhuhur)%24)
          m = ((zhuhur*60)%60)+2
          s = (zhuhur*3600)%60
          Dzuhur = "%d:%02d:%02d"%(h,m,s)
      #ASHAR
          KA = 1 # syafii
          h_AS=(math.atan(1/(KA+math.tan((abs(Delta-L))*(math.pi/180)))))*180/math.pi
          ASHAR = (math.sin(math.radians(h_AS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
          HA_AS = math.degrees(math.acos(ASHAR))
          "waktu ashar"
          ashar = Transit + (HA_AS)/15
          h = (int(ashar)%24)
          m = ((ashar*60)%60)
          s = (ashar*3600)%60
          Ashar = "%d:%02d:%02d"%(h,m,s)
      #Magrib
          h_MG = -0.8333-00.0347*math.sqrt(H)
          MAGRIB = (math.sin(math.radians(h_MG))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
          HA_MG = math.degrees(math.acos(MAGRIB))
          "waktu magrib"
          magrib = Transit + (HA_MG)/15
          h = (int(magrib)%24)
          m = ((magrib*60)%60)
          s = (magrib*3600)%60
          Magrib = "%d:%02d:%02d"%(h,m,s)
      #Isya
          h_IS = -18
          ISYA = (math.sin(math.radians(h_IS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
          HA_IS = math.degrees(math.acos(ISYA))
          "waktu isya"
          isya = Transit + (HA_IS)/15
          h = (int(isya)%24)
          m = ((isya*60)%60)
          s = (isya*3600)%60
          Isya = "%d:%02d:%02d"%(h,m,s)
      #Shubuh
          h_SH = -20
          SHUBUH = (math.sin(math.radians(h_SH))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
          HA_SH = math.degrees(math.acos(SHUBUH))
          "waktu shubuh"
          shubuh = Transit - (HA_SH)/15
          h = (int(shubuh)%24)
          m = ((shubuh*60)%60)
          s = (shubuh*3600)%60
          Shubuh = "%d:%02d:%02d"%(h,m,s)
      #Terbit
          "waktu terbit"
          terbit = Transit - (HA_MG)/15
          h = (int(terbit)%24)
          m = ((terbit*60)%60)
          s = (terbit*3600)%60
          Wter = "%d:%02d:%02d"%(h,m,s)
          jadwal1 = f"Jadwal Shalat \n {lokasikota}, {spinbox1.get()} {spinbox2.get()} {spinbox3.get()}"  f"\nShubuh : {Shubuh}"  f"\nTerbit : {Wter}" f"\nDzuhur : {Dzuhur}" f"\nAshar : {Ashar}" f"\nMagrib : {Magrib}" f"\nIsya : {Isya}"
          teks.configure(text=jadwal1)
          for i in range(0,n):
            i = i+1
            month = sum + i
            total = total_day + month
            jd = 1948438.5 + total
            JD_lokal = jd - Z/24
            T = 2*3.14159265359*(JD_lokal-2451545)/365.25
            Delta = 0.37877 + 23.264*math.sin(math.radians(57.297*T-79.547)) + 0.3812*math.sin(math.radians(2*57.297*T-82.682)) + 0.17132*math.sin(math.radians (3*57.297*T-59.722))
            U = (JD_lokal - 2451545)/36525
            L0x = 280.46607 + 36000.7698*U
            L0 = L0x%360
            rad = np.pi/180
            et = ((-(1789+237*U))*(np.sin(L0*rad)))-((7146-62*U)*(np.cos(L0*rad)))+((9934-14*U)*(np.sin(2*L0*rad)))-((29+5*U)*(np.cos(2*L0*rad)))+((74+10*U)*(np.sin(3*L0*rad)))+((320-4*U)*(np.cos(3*L0*rad)))-((212*np.sin(4*L0*rad)))
            ET = et/1000
            Transit = 12 + Z -(BU/15) - (ET/60)
            # Dzuhur
            zhuhur = Transit # +2menit
            h = (int(zhuhur)%24)
            m = ((zhuhur*60)%60)+2
            s = (zhuhur*3600)%60
            Dzuhur = "%d:%02d:%02d"%(h,m,s)
            #ASHAR
            KA = 1 # syafii
            h_AS=(math.atan(1/(KA+math.tan((abs(Delta-L))*(math.pi/180)))))*180/math.pi
            ASHAR = (math.sin(math.radians(h_AS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_AS = math.degrees(math.acos(ASHAR))
            "waktu ashar"
            ashar = Transit + (HA_AS)/15
            h = (int(ashar)%24)
            m = ((ashar*60)%60)
            s = (ashar*3600)%60
            Ashar = "%d:%02d:%02d"%(h,m,s)
            #Magrib
            h_MG = -0.8333-00.0347*math.sqrt(H)
            MAGRIB = (math.sin(math.radians(h_MG))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_MG = math.degrees(math.acos(MAGRIB))
            "waktu magrib"
            magrib = Transit + (HA_MG)/15
            h = (int(magrib)%24)
            m = ((magrib*60)%60)
            s = (magrib*3600)%60
            Magrib = "%d:%02d:%02d"%(h,m,s)
            #Isya
            h_IS = -18
            ISYA = (math.sin(math.radians(h_IS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_IS = math.degrees(math.acos(ISYA))
            "waktu isya"
            isya = Transit + (HA_IS)/15
            h = (int(isya)%24)
            m = ((isya*60)%60)
            s = (isya*3600)%60
            Isya = "%d:%02d:%02d"%(h,m,s)
            #Shubuh
            h_SH = -20
            SHUBUH = (math.sin(math.radians(h_SH))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_SH = math.degrees(math.acos(SHUBUH))
            "waktu shubuh"
            shubuh = Transit - (HA_SH)/15
            h = (int(shubuh)%24)
            m = ((shubuh*60)%60)
            s = (shubuh*3600)%60
            Shubuh = "%d:%02d:%02d"%(h,m,s)
            #Terbit
            "waktu terbit"
            terbit = Transit - (HA_MG)/15
            h = (int(terbit)%24)
            m = ((terbit*60)%60)
            s = (terbit*3600)%60
            Wter = "%d:%02d:%02d"%(h,m,s)
            data = f"{i} {spinbox2.get()}"
            my_tree.insert(parent='',index='end', text= data, values= (f"{Shubuh}",f"{Wter}" ,f"{Dzuhur}" ,f"{Ashar}" ,f"{Magrib}" ,f"{Isya}"))
      tombol1 = ttk.Button(root,text="Cari", command=tekan)
      tombol1.place(x = 330, y =120)
      teks = Label(f2,text="")
      teks.place(x=500, y=20)
      
      excel_path = r".\Jadwal 1 Bulan.xlsx"
      file = pathlib.Path("Jadwal 1 Bulan.xlsx")
      if file.exists():
        pass
      else : 
        file = Workbook()
        sheet = file.active
        sheet['A1'] = "Tanggal"
        sheet['B1'] = "Shubuh"
        sheet['C1'] = "Terbit"
        sheet['D1'] = "Dzuhur"
        sheet['E1'] = "Ashar"
        sheet['F1'] = "Magrib"
        sheet['G1'] = "Isya"
        sheet['I1'] = "Lokasi"
        file.save("Jadwal 1 Bulan.xlsx")

      def download() :
        lokasikota = cb.get()
        D = int(spinbox1.get())
        M = spinbox2.get()
        M1 = spinbox2.get()
        Y = int (spinbox3.get())
        if (cb.get() == "Jakarta") :
          L = csv_data['Lintang'].iloc[0]
          BU = csv_data['Bujur'].iloc[0]
          Z = csv_data['Zona Waktu'].iloc[0]
          H = csv_data['Ketinggian'].iloc[0]
        elif (cb.get() ==  "Mojokerto") :
          L = csv_data['Lintang'].iloc[1]
          BU = csv_data['Bujur'].iloc[1]
          Z = csv_data['Zona Waktu'].iloc[1]
          H = csv_data['Ketinggian'].iloc[1]
        elif (cb.get() == "Manado"):
          L = csv_data['Lintang'].iloc[2]
          BU = csv_data['Bujur'].iloc[2]
          Z = csv_data['Zona Waktu'].iloc[2]
          H = csv_data['Ketinggian'].iloc[2]
        elif (cb.get() ==  "Medan"):
          L = csv_data['Lintang'].iloc[3]
          BU = csv_data['Bujur'].iloc[3]
          Z = csv_data['Zona Waktu'].iloc[3]
          H = csv_data['Ketinggian'].iloc[3]
        elif (cb.get() ==  "Natuna") :
          L = csv_data['Lintang'].iloc[4]
          BU = csv_data['Bujur'].iloc[4]
          Z = csv_data['Zona Waktu'].iloc[4]
          H = csv_data['Ketinggian'].iloc[4]
        elif (cb.get() == "Malang") :
          L = csv_data['Lintang'].iloc[5]
          BU = csv_data['Bujur'].iloc[5]
          Z = csv_data['Zona Waktu'].iloc[5]
          H = csv_data['Ketinggian'].iloc[5]
        elif (cb.get() ==  "Makasar") :
          L = csv_data['Lintang'].iloc[6]
          BU = csv_data['Bujur'].iloc[6]
          Z = csv_data['Zona Waktu'].iloc[6]
          H = csv_data['Ketinggian'].iloc[6]
        elif (cb.get() ==  "Bogor") :
          L = csv_data['Lintang'].iloc[7]
          BU = csv_data['Bujur'].iloc[7]
          Z = csv_data['Zona Waktu'].iloc[7]
          H = csv_data['Ketinggian'].iloc[7]
        elif (cb.get() ==  "Denpasar") :
          L = csv_data['Lintang'].iloc[8]
          BU = csv_data['Bujur'].iloc[8]
          Z = csv_data['Zona Waktu'].iloc[8]
          H = csv_data['Ketinggian'].iloc[8]
        elif (cb.get() ==  "Gorontalo") :
          L = csv_data['Lintang'].iloc[9]
          BU = csv_data['Bujur'].iloc[9]
          Z = csv_data['Zona Waktu'].iloc[9]
          H = csv_data['Ketinggian'].iloc[9] 
        
        if (M =="Januari") :
         M = 1
         n = 31
        elif (M =="Februari") :
          M = 2
          if Y%4 == 0 :
            n = 29
          else :
            n = 28
        elif (M =="Maret") :
          M = 3
          n =31
        elif (M =="April") :
          M = 4
          n = 30
        elif (M =="Mei") :
          M = 5
          n = 31
        elif (M=="Juni") :
          M = 6
          n =30
        elif (M=="Juli") :
          M = 7
          n = 31
        elif (M=="Agustus") :
          M = 8
          n =31
        elif (M=="September") :
          M = 9
          n = 30
        elif (M=="Oktober") :
          M = 10
          n = 31
        elif (M=="November") :
          M = 11
          n = 30
        elif (M =="Desember") :
          M = 12
          n = 31

        if (M=="Muharram") :
          M = M1
          M1 = 1
          n = 30
        elif (M1=="Shafar") :
          M = M
          M1 = 2
          n = 29
        elif (M1=="Rabi'ul Awal") :
          M = M
          M1 = 3
          n = 30
        elif (M1=="Rabi'ul akhir") :
          M = M
          M1 = 4
          n = 29
        elif (M1=="Jumadil Awal") :
          M = M1
          M1 = 5
          n = 30
        elif (M1=="Jumadil Akhir") :
          M = M1
          M1 = 6
          n =29
        elif (M1=="Rajab") :
          M = M1
          M1 = 7
          n = 30
        elif (M1=="Sya'ban") :
          M = M1
          M1 = 8
          n = 29
        elif (M1=="Ramadhan") :
          M = M1
          M1= 9
          n = 30
        elif (M1=="Syawwal") :
          M = M1
          M1 = 10
          n = 29
        elif (M1=="Zulqaidah") :
          M = M1
          M1 = 11
          n = 30
        elif (M1=="Zulhijjah") :
          M = M1
          M1 = 12
          if Y%30 == 2 :
            n = 30
          elif Y%30 == 5 :
            n = 30
          elif Y%30 == 7 :
            n = 30
          elif Y%30 == 10 :
            n = 30
          elif Y%30 == 13 :
            n = 30
          elif Y%30 == 16 :
            n = 30
          elif Y%30 == 18 :
            n = 30
          elif Y%30 == 21 :
            n = 30
          elif Y%30 == 24 :
            n = 30
          elif Y%30 == 26 :
            n = 30
          elif Y%30 == 29 :
            n = 30
          else :
            n =29 
        
        if (c1.get()==1) & (c2.get()==0) :
          if M > 2 :
            M = M
            Y = Y
          elif M in [1 or 2]:
            M = M +12
            Y = Y -1
          #masehi
          A  = int(Y/100)
          B = 2 + int(A/4)-A
          JD = 1720994.5 + int(365.25*Y) + int(30.6001*(M+1))+ B + D +0.5
          JD_lokal = JD - Z/24
          for i in range(0,n):
            i = i+1
            JD = 1720994.5 + int(365.25*Y) + int(30.6001*(M+1))+ B + i +0.5
            JD_lokal = JD - Z/24
            T = 2*3.14159265359*(JD_lokal-2451545)/365.25
            Delta = 0.37877 + 23.264*math.sin(math.radians(57.297*T-79.547)) + 0.3812*math.sin(math.radians(2*57.297*T-82.682)) + 0.17132*math.sin(math.radians (3*57.297*T-59.722))
            U = (JD_lokal - 2451545)/36525
            L0x = 280.46607 + 36000.7698*U
            L0 = L0x%360
            rad = np.pi/180
            et = ((-(1789+237*U))*(np.sin(L0*rad)))-((7146-62*U)*(np.cos(L0*rad)))+((9934-14*U)*(np.sin(2*L0*rad)))-((29+5*U)*(np.cos(2*L0*rad)))+((74+10*U)*(np.sin(3*L0*rad)))+((320-4*U)*(np.cos(3*L0*rad)))-((212*np.sin(4*L0*rad)))
            ET = et/1000
            Transit = 12 + Z -(BU/15) - (ET/60)  
            # Dzuhur
            zhuhur = Transit # +2menit
            h = (int(zhuhur)%24)
            m = ((zhuhur*60)%60)+2
            s = (zhuhur*3600)%60
            Dzuhur = "%d:%02d:%02d"%(h,m,s)
            #ASHAR
            KA = 1 # syafii
            h_AS=(math.atan(1/(KA+math.tan((abs(Delta-L))*(math.pi/180)))))*180/math.pi
            ASHAR = (math.sin(math.radians(h_AS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_AS = math.degrees(math.acos(ASHAR))
            "waktu ashar"
            ashar = Transit + (HA_AS)/15
            h = (int(ashar)%24)
            m = ((ashar*60)%60)
            s = (ashar*3600)%60
            Ashar = "%d:%02d:%02d"%(h,m,s)
            #Magrib
            h_MG = -0.8333-00.0347*math.sqrt(H)
            MAGRIB = (math.sin(math.radians(h_MG))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_MG = math.degrees(math.acos(MAGRIB))
            "waktu magrib"
            magrib = Transit + (HA_MG)/15
            h = (int(magrib)%24)
            m = ((magrib*60)%60)
            s = (magrib*3600)%60
            Magrib = "%d:%02d:%02d"%(h,m,s)
            #Isya
            h_IS = -18
            ISYA = (math.sin(math.radians(h_IS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_IS = math.degrees(math.acos(ISYA))
            "waktu isya"
            isya = Transit + (HA_IS)/15
            h = (int(isya)%24)
            m = ((isya*60)%60)
            s = (isya*3600)%60
            Isya = "%d:%02d:%02d"%(h,m,s)
            #Shubuh
            h_SH = -20
            SHUBUH = (math.sin(math.radians(h_SH))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_SH = math.degrees(math.acos(SHUBUH))
            "waktu shubuh"
            shubuh = Transit - (HA_SH)/15
            h = (int(shubuh)%24)
            m = ((shubuh*60)%60)
            s = (shubuh*3600)%60
            Shubuh = "%d:%02d:%02d"%(h,m,s)
            #Terbit
            "waktu terbit"
            terbit = Transit - (HA_MG)/15
            h = (int(terbit)%24)
            m = ((terbit*60)%60)
            s = (terbit*3600)%60
            Wter = "%d:%02d:%02d"%(h,m,s)
            data = f"{i} {spinbox2.get()}"
            tanggal_value = data
            shubuh_value = f"{Shubuh}"
            Wter_value = f"{Wter}"
            Dzuhur_value = f"{Dzuhur}"
            Ashar_value = f"{Ashar}"
            Magrib_value = f"{Magrib}"
            Isya_value = f"{Isya}"
            lokasi_value = f"{cb.get()},{spinbox3.get()}"
            file=openpyxl.load_workbook("Jadwal 1 Bulan.xlsx")
            sheet = file.active
            sheet.cell(column=1, row=sheet.max_row+1, value=tanggal_value)
            sheet.cell(column=2, row=sheet.max_row, value=shubuh_value)
            sheet.cell(column=3, row=sheet.max_row, value=Wter_value)
            sheet.cell(column=4, row=sheet.max_row, value=Dzuhur_value)
            sheet.cell(column=5, row=sheet.max_row, value=Ashar_value)
            sheet.cell(column=6, row=sheet.max_row, value=Magrib_value)
            sheet.cell(column=7, row=sheet.max_row, value=Isya_value)
            sheet.cell(column=9, row=1, value=lokasi_value)
            file.save("Jadwal 1 Bulan.xlsx")
          messagebox.showinfo(title="Sukses", message= "Jadwal 1 Bulan \nBerhasil Di download") 
        elif  (c1.get()==0) & (c2.get()==1) :
          leap_year = [2,5,7,10,13,16,18,21,24,26,29]
          year2 = (Y-1)%30
          if year2 == 2:
            length = len(leap_year[0:1])
          elif 5 <= year2 <7 :
            length = len(leap_year[0:2])
          elif 7 <= year2 < 10 :
            length = len(leap_year[0:3])
          elif 10 <= year2 < 13 :
            length = len(leap_year[0:4])
          elif 13 <= year2 < 16:
            length = len(leap_year[0:5])
          elif 16<= year2 < 18 :
            length = len(leap_year[0:6])
          elif 18<= year2 < 21 :
            length = len(leap_year[0:7])
          elif 21<= year2 < 24 :
            length = len(leap_year[0:8])
          elif 24 <= year2 < 26 :
            length = len(leap_year[0:9])
          elif 26 <= year2 < 29:
            length = len(leap_year[0:10])
          elif 29 <= year2 >= 30 :
            length = len(leap_year[0:11])
          else : 
            length = 0
          year = int((Y - 1)/30)
          total_day = (year*(354*30+11)) + (year2*354 + length)
          hijriah = { 1: 30, 2 : 29, 3 : 30, 4 : 29, 5: 30, 6 : 29, 7 : 30, 8 : 29, 9 : 30, 10 : 29, 11 : 30, 12 : 29}
          key = list(hijriah.keys())
          val = list(hijriah.values())
      
          if M1 == 2:
            sum = val[0]
          elif M1 == 3 :
            sum = val[0]+val[1]
          elif M1 == 4 :
            sum = val[0]+val[1]+val[2]
          elif M1 == 5 :
            sum = val[0]+val[1]+val[2]+val[3]
          elif M1 == 6 :
            sum = val[0]+val[1]+val[2]+val[3]+val[4]
          elif M1 == 7 :
            sum = val[0]+val[1]+val[2]+val[3]+val[4]+val[5]
          elif M1 == 8 :
            sum = val[0]+val[1]+val[2]+val[3]+val[4]+val[5]+val[6]
          elif M1 == 9 :
            sum = val[0]+val[1]+val[2]+val[3]+val[4]+val[5]+val[6]+val[7]
          elif M1 == 10 : 
            sum = val[0]+val[1]+val[1]+val[2]+val[3]+val[4]+val[8]+val[9]+val[10]
          elif M1 == 11 :
            sum = val[0]+val[1]+val[1]+val[2]+val[3]+val[4]+val[8]+val[9]+val[10]+val[11]
          elif M1 == 12 : 
            sum = val[0]+val[1]+val[1]+val[2]+val[3]+val[4]+val[8]+val[9]+val[10]+val[11]+val[11]
          else:
            sum = 0
          month = sum + D
          total = total_day + month
          jd = 1948438.5 + total
          JD_lokal = jd - Z/24
          for i in range(0,n):
            i = i+1
            month = sum + i
            total = total_day + month
            jd = 1948438.5 + total
            JD_lokal = jd - Z/24
            T = 2*3.14159265359*(JD_lokal-2451545)/365.25
            Delta = 0.37877 + 23.264*math.sin(math.radians(57.297*T-79.547)) + 0.3812*math.sin(math.radians(2*57.297*T-82.682)) + 0.17132*math.sin(math.radians (3*57.297*T-59.722))
            U = (JD_lokal - 2451545)/36525
            L0x = 280.46607 + 36000.7698*U
            L0 = L0x%360
            rad = np.pi/180
            et = ((-(1789+237*U))*(np.sin(L0*rad)))-((7146-62*U)*(np.cos(L0*rad)))+((9934-14*U)*(np.sin(2*L0*rad)))-((29+5*U)*(np.cos(2*L0*rad)))+((74+10*U)*(np.sin(3*L0*rad)))+((320-4*U)*(np.cos(3*L0*rad)))-((212*np.sin(4*L0*rad)))
            ET = et/1000
            Transit = 12 + Z -(BU/15) - (ET/60)  
            # Dzuhur
            zhuhur = Transit # +2menit
            h = (int(zhuhur)%24)
            m = ((zhuhur*60)%60)+2
            s = (zhuhur*3600)%60
            Dzuhur = "%d:%02d:%02d"%(h,m,s)
            #ASHAR
            KA = 1 # syafii
            h_AS=(math.atan(1/(KA+math.tan((abs(Delta-L))*(math.pi/180)))))*180/math.pi
            ASHAR = (math.sin(math.radians(h_AS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_AS = math.degrees(math.acos(ASHAR))
            "waktu ashar"
            ashar = Transit + (HA_AS)/15
            h = (int(ashar)%24)
            m = ((ashar*60)%60)
            s = (ashar*3600)%60
            Ashar = "%d:%02d:%02d"%(h,m,s)
            #Magrib
            h_MG = -0.8333-00.0347*math.sqrt(H)
            MAGRIB = (math.sin(math.radians(h_MG))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_MG = math.degrees(math.acos(MAGRIB))
            "waktu magrib"
            magrib = Transit + (HA_MG)/15
            h = (int(magrib)%24)
            m = ((magrib*60)%60)
            s = (magrib*3600)%60
            Magrib = "%d:%02d:%02d"%(h,m,s)
            #Isya
            h_IS = -18
            ISYA = (math.sin(math.radians(h_IS))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_IS = math.degrees(math.acos(ISYA))
            "waktu isya"
            isya = Transit + (HA_IS)/15
            h = (int(isya)%24)
            m = ((isya*60)%60)
            s = (isya*3600)%60
            Isya = "%d:%02d:%02d"%(h,m,s)
            #Shubuh
            h_SH = -20
            SHUBUH = (math.sin(math.radians(h_SH))-((math.sin(math.radians(L))*(math.sin(math.radians( Delta))))))/((math.cos(math.radians(L))*(math.cos(math.radians( Delta)))))
            HA_SH = math.degrees(math.acos(SHUBUH))
            "waktu shubuh"
            shubuh = Transit - (HA_SH)/15
            h = (int(shubuh)%24)
            m = ((shubuh*60)%60)
            s = (shubuh*3600)%60
            Shubuh = "%d:%02d:%02d"%(h,m,s)
            #Terbit
            "waktu terbit"
            terbit = Transit - (HA_MG)/15
            h = (int(terbit)%24)
            m = ((terbit*60)%60)
            s = (terbit*3600)%60
            Wter = "%d:%02d:%02d"%(h,m,s)
            data = f"{i} {spinbox2.get()}"
            tanggal_value = data
            shubuh_value = f"{Shubuh}"
            Wter_value = f"{Wter}"
            Dzuhur_value = f"{Dzuhur}"
            Ashar_value = f"{Ashar}"
            Magrib_value = f"{Magrib}"
            Isya_value = f"{Isya}"
            lokasi_value = f"{cb.get()},{spinbox3.get()}"
            file=openpyxl.load_workbook("Jadwal 1 Bulan.xlsx")
            sheet = file.active
            sheet.cell(column=1, row=sheet.max_row+1, value=tanggal_value)
            sheet.cell(column=2, row=sheet.max_row, value=shubuh_value)
            sheet.cell(column=3, row=sheet.max_row, value=Wter_value)
            sheet.cell(column=4, row=sheet.max_row, value=Dzuhur_value)
            sheet.cell(column=5, row=sheet.max_row, value=Ashar_value)
            sheet.cell(column=6, row=sheet.max_row, value=Magrib_value)
            sheet.cell(column=7, row=sheet.max_row, value=Isya_value)
            sheet.cell(column=9, row=1, value=lokasi_value)
            file.save("Jadwal 1 Bulan.xlsx")
          messagebox.showinfo(title="Sukses", message= "Jadwal 1 Bulan \nBerhasil Di download")  
      tomboldo = ttk.Button(root, text ="download", command=download)
      tomboldo.place( x = 330, y =150)
    c1 =IntVar()
    c2 = IntVar()
    CB = Checkbutton(root,text="Masehi",command=a, variable=c1)
    CB. place(x=250, y = 120)
    CB1 = Checkbutton(root, text="Hijriyah",command=a, variable=c2)
    CB1.place(x=250, y = 150)

    scrollbarx= Scrollbar(f2,orient=HORIZONTAL)
    scrollbary = Scrollbar(f2, orient=VERTICAL)
    my_tree = ttk.Treeview(f2)
    my_tree.place(relx=0.04, rely=0.3, width=630, height=310)
    my_tree.configure(yscrollcommand=scrollbary.set, xscrollcommand=scrollbarx.set)
    my_tree.configure(selectmode="extended")
    
    scrollbary.configure(command=my_tree.yview)
    scrollbarx.configure(command=my_tree.xview)
    scrollbary.place(relx=0.939, rely=0.30, width=22, height=310)
    scrollbarx.place(relx=0.04, rely=0.864, width=630, height=22)
    
    my_tree.configure(
      columns= (
        "Shubuh",
        "Terbit",
        "Dzuhur",
        "Ashar",
        "Magrib",
        "Isya"
      )
    )    
    my_tree.heading("#0", text="Tanggal",anchor=W)
    my_tree.heading("Shubuh", text="Shubuh",anchor=W)
    my_tree.heading("Terbit", text="Terbit",anchor=W)
    my_tree.heading("Dzuhur", text="Dzuhur",anchor=W)
    my_tree.heading("Ashar", text="Ashar",anchor=W)
    my_tree.heading("Magrib", text="Magrib",anchor=W)
    my_tree.heading("Isya", text="Isya",anchor=W)
    my_tree.column("#0", stretch=NO, minwidth=25, width=125)
    my_tree.column("#1", stretch=NO, minwidth=0, width=200)
    my_tree.column("#2", stretch=NO, minwidth=0, width=160)
    my_tree.column("#3", stretch=NO, minwidth=0, width=160)
    my_tree.column("#4", stretch=NO, minwidth=0, width=160)
    my_tree.column("#5", stretch=NO, minwidth=0, width=160)
    my_tree.column("#6", stretch=NO, minwidth=25, width=160)  
    toggle_win()

def bantuan():
    f1.destroy()
    f2=Frame(root,width=700,height=550,bg='#262626')
    f2.place(x=0,y=45)
    l=Label(f2,text='Informasi lebih lanjut?\nEmail : tafahz@student.ub.ac.id',fg='white',bg='#262626')
    l.config(font=('Comic Sans MS',12))
    l.place(x=220,y=350)
    l1=Label(f2,text='Dalam mencari jadwal shalat yang diinginkan yang perlu anda lakukan ialah :',fg='white',bg='#262626')
    l1.config(font=('Comic Sans MS',14))
    l1.place(x=20,y=50)
    hm=Label(f2,text='1. Pilih menu home\n2.Pilih lokasi kota\n3Pilih Kalender\n4.Masukkan tanggal\n5. Masukkan bulan\n6.Masukkan tahun \n7.Tekan tombol "Cari"\n8. Tekan tombol "download" untuk mendapatkan\njadwal 1 bulan dalam bentuk excel',fg='white',bg='#262626')
    hm.config(font=('Comic Sans MS',14))
    hm.place(x=140,y=80)
    toggle_win()
  
def definisi():
    f1.destroy()
    f2=Frame(root,width=700,height=550,bg='#262626')
    f2.place(x=0,y=45)
    l=Label(f2,text='Program Waktu Shalat',fg='white',bg='#262626')
    l.config(font=('Comic Sans MS',40))
    l.place(x=80,y=50)
    l1=Label(f2,text='Tugas Mata Kuliah Komputasi Astronomi',fg='white',bg='#262626')
    l1.config(font=('Comic Sans MS',14))
    l1.place(x=170,y=130)
    df=Label(f2,text='Dosen Pengampu : Dr.rer.nat. Abdurrouf, S.Si., M.Si\n Nama : Tafa Hafidzah \nNIM : 205090301111013',fg='white',bg='#262626')
    df.config(font=('Comic Sans MS',12))
    df.place(x=150,y=190)
    

root.mainloop()
