import os
import pandas as pd
import shutil
import stat

#File location for Drawings
#DRW_source = r"C:\LAC Conveyors\Playground\Playground Kon\Design Table Trial\Drawing Detail\\"
DRW_source = r"C:\LAC Conveyors\Playground\Playground Kon\Design Table Trial\Macro DRW\\"
#File location for BOMs
XLS_source = r"C:\LAC Conveyors\Playground\Playground Kon\Design Table Trial\Drawing Detail\\"
#File location for library
End_source = r"C:\LAC Conveyors\Playground\Playground Chris Yates\Python Tests\\"

#------------------------------------------------------------ 
def CopyFiles(source_folder, destination_folder, bom_folder):
    prtName = FindExcelDocs(bom_folder)
    for file in os.listdir(source_folder):
        source = source_folder + file
        for BOM in prtName:
            num = prtName.index(BOM)
            dest = destination_folder[:-2] + "\\" + File_name[num] + "\\"
            #destination = dest + file
            CreateSubFolders(dest)
            #destination = dest + file
            for item in BOM:
                sort(item, source, dest, file, 0)
                sort(item, source, dest, file, 1)
                sort(item, source, dest, file, 2)
                sort(item, source, dest, file, 3)
    
    CopyAssy(source_folder, destination_folder)
    #print(File_name)
#------------------------------------------------------------ 
def FindExcelDocs(source):
    global File_name
    BOMs = []
    File_name = []
    for file in os.listdir(source):
        if file.endswith(".xls") or file.endswith(".xlsx"):
            CreateFolder(file, End_source)
            File_name.append(dirName)
            df = pd.read_excel(source + file, usecols="B")
            BOM = []
            for num in range(len(df)):
                BOM.append(df.iloc[num][0])
            BOMs.append(BOM)
    return(BOMs)
#------------------------------------------------------------                    
def CreateFolder(name, location):   
    global dirName
    try:
        dirName = name.strip(".xlsx")
    except:
        dirName = name.strip(".xls")
        
    path = os.path.join(location, dirName)
    
    try:
        os.mkdir(path)
        print("Folder '% s' created" % dirName)
    except OSError as error:
        pass
#------------------------------------------------------------         
def CopyAssy(source_folder, destination_folder):
    for num in File_name:
        i = File_name.index(num)
        for file in os.listdir(source_folder):            
            source = source_folder + file
            dest = destination_folder[:-2] + "\\" + File_name[i] + "\\"
            destination = dest + file
            Assy = str(File_name[i]) + ".SLDDRW"
            if os.path.isfile(source) and Assy == file:
                try:
                    #os.chmod(destination, stat.S_IWRITE)
                    shutil.copy(source, destination)
                    print("Copied assembly", file)
                    #continue
                except:
                    os.chmod(destination, stat.S_IWRITE)
                    shutil.copy(source, destination)
                    print("Updated assembly", file)
                    #continue
#------------------------------------------------------------         
def CreateSubFolders(location):
    Type = ["Fabrications", "Lasered", "Machined", "Rollers"]
    for i in Type:
        path = os.path.join(location, i)   
        try:
            os.mkdir(path)
            print("Folder '% s' created" % i)
        except OSError as error:
            pass
#------------------------------------------------------------ 
def sort(item, source, dest, file, option):
    #destination = dest + file
    Types = [".SLDDRW", ".pdf", ".DXF"]
    for _type in Types:
        i = ["L", "M", "F", "R"]
        Folders = ["Lasered\\", "Machined\\", "Fabrications\\", "Rollers\\"]
        doc = str(item) + _type
        if os.path.isfile(source) and doc == file:
            if item[-1] == i[option]:
                destination = dest + Folders[option] + file
                try:
                    os.chmod(destination, stat.S_IWRITE)
                    shutil.copy(source, destination)
                    print("Updated ", file)
                except:
                    shutil.copy(source, destination)
                    print("Copied ", file)
#------------------------------------------------------------

CopyFiles(DRW_source, End_source, XLS_source)
