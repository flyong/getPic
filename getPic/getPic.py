# -*- coding: utf-8 -*-
import os,shutil,zipfile,tkinter,re
from tkinter import filedialog
from win32com import client as wc
    

def doc2docx(path):#doc2docx,?????µ?docx??·??
    path_docx=[]
    word = wc.Dispatch("Word.Application")
    for file in path:
        if re.match('doc$',file):#????doc?ļ???????Ϊdocx
            doc = word.Documents.Open(file)
            newPath=file[:-4]+'.docx'
            doc.SaveAs(newPath, 12) #????Ϊdocx
            doc.Close()
            path_docx.append(newPath)
        else:
            path_docx.append(file)
    word.Quit()
    return path_docx    

def getPic(path_docx):#??docx?ļ?תΪzip    
    try:
        for docx in path_docx:
            docxName=os.path.split(docx)[-1]
            folder=os.path.dirname(docx) 
            storeFolder=folder+'\\'+docxName                   
            tempFolder=storeFolder+'temp'
            os.makedirs(tempFolder)#?????ļ??????ڴ?????ѹ?ļ?
            zip=docx[:-5]+r'.zip'
            os.rename(docx,zip)#docx?ļ???????Ϊzip 
            f=zipfile.ZipFile(zip,'r')#??ȡzip  
            for file in f.namelist():                
                f.extract(file,tempFolder)#??ѹ??ָ???ļ???            
            f.close()     
            pic=os.listdir(os.path.join(tempFolder,'word/media')) 
            for i in pic:
                picName=docxName+'_'+i
                shutil.copy(os.path.join(tempFolder+'word/media',i),os.path.join(storeFolder, picName))#???Ƶ?ָ???ļ???
            for i in os.listdir(tempFolder):
                if os.path.isdir(os.path.join(tempFolder),i):#ɾ????ת?ļ???
                    shutil.rmtree(os.path.join(tempFolder,i))
    
path=tkinter.filedialog.askopenfilenames()
path_docx=doc2docx(path)
getPic(path_docx)



        
