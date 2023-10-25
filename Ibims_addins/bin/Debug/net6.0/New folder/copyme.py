import os,os.path
import shutil
from tkinter import messagebox

curentDir= os.path.abspath(os.path.curdir)
destination = os.path.join(os.path.expanduser('~'),"AppData\\Roaming\\Autodesk\\Revit\\Addins")
dllPath = os.path.join(curentDir,"IBIMSGen.dll")
addinPath = os.path.join(curentDir,"IBIMSGen.addin")
message=""
for directory in os.listdir(destination):
    if(os.path.isdir(os.path.join(destination,directory))):
        toDll = os.path.join(destination,directory,"IBIMSGen.dll")
        toAddin = os.path.join(destination,directory,"IBIMSGen.addin")
        try:
            shutil.copy(dllPath,toDll)
            shutil.copy(addinPath,toAddin)
            message+='Copied to v. '+directory+" successfully.\n"
        except:
            message+="====================="
            message+='\nfile not copied to V.'+directory+'\nmaybe revit is open, Close the revit and try again\n'
            message+="=====================\n"

message+="\n\n\n\nBy: Omar Elshaf3y | 2023 :)\n    m.me\\o.elshaf3y"
messagebox.showinfo("Ibims",message)