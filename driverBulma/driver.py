
import  wmi 
import json
import os
import tkinter as tk
from tkinter import ttk


import webbrowser

def listDrivers(saveToFile=False):
    try:
        computer=wmi.WMI()
        drivers=[]
        for device in computer.Win32_PnPEntity():
            if device.Name and device.DeviceID:
                status="Yüklü" if device.Status == "OK" else "Yüklü Değil"
                driver_info={
                    "Cihaz":device.Name,
                    "Device ID": device.DeviceID,
                    "Üretici": device.Manufacturer if hasattr(device,'Manufacturer') else "Bilinmiyor",
                    "Durum": status,
                    "Sürücü Linki": "Sürücüyü Araştır"
                }
                drivers.append(driver_info)
        if saveToFile:
            filePath=os.path.join(os.getcwd(),"drivers_info.json")
            with open(filePath,"w",encoding="utf-8")as file:
                json.dump(drivers,file,ensure_ascii=False, indent=4)
            print(f"Sürücü Bilgileri {filePath} dosyasına kayıt edildi")
        return drivers
    except Exception as e:
        print(f"Hata : {str(e)}")
        return[]

def displayDriversinTable():
    drivers=listDrivers()
    root=tk.Tk()
    root.title("Bilgisayar Sürücüleri")
    root.geometry("1000x500")

    searchFrame=tk.Frame(root)
    searchFrame.pack(fill="x",padx=10,pady=10)

    searchLabel=tk.Label(searchFrame,text="Ara:")
    searchLabel.pack(side="left",padx=5)
    
    searchEntry=tk.Entry(searchFrame)
    searchEntry.pack(side="left",fill="x",expand=True,padx=5)

    columns=("Cihaz","DeviceID","Üretici","Durum","Sürücü Linki")
    tree=ttk.Treeview(root,columns=columns,show="headings")

    for col in columns:
        tree.heading(col,text=col)
        tree.column(col,width=200)
    
    def populateTreewiew(data):
        for item in tree.get_children():
            tree.delete(item)
        for driver in data :
            item_id=tree.insert("",tk.END,values=(
                driver.get("Cihaz","Bilinmiyor"),
                driver.get("DeviceID","Bilinmiyor"),
                driver.get("Üretici","Bilinmiyor"),
                driver.get("Durum","Bilinmiyor"),
                "Sürücüyü Araştır"
            ))

            if driver.get("Durum")=="Yüklü":
                tree.item(item_id,tags=("yuklu",))
            else:
                tree.item(item_id,tags=("yuklu_degil",))

    tree.tag_configure("yuklu", background="lightgreen")
    tree.tag_configure("yuklu_degil", background="lightcoral")

    populateTreewiew(drivers)
    tree.pack(expand=True, fill="both")


    def searchDrivers(event=None):
        query=searchEntry.get().lower()
        filteredDrivers=[

            driver for driver in drivers
            if
             (driver.get("Cihaz") and query in driver["Cihaz"].lower()) or
             (driver.get("DeviceID") and query in driver["DeviceID"].lower()) or
             (driver.get("Üretici") and query in driver["Üretici"].lower()) or
             (driver.get("Durum") and query in driver["Durum"].lower()) 
        ]
        populateTreewiew(filteredDrivers)
    searchEntry.bind("<KeyRelease>",searchDrivers)

    def onTreeviewDoubleClick(event):
        item=tree.selection()[0]
        deviceName=tree.item(item,"values")[0]
        driverLink=f"https://www.google.com/search?q={deviceName} driver"
        if driverLink:
            webbrowser.open(driverLink)

    tree.bind("<Double-1>", onTreeviewDoubleClick)
    root.mainloop()

if __name__=="__main__":
    displayDriversinTable()
