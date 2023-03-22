import os
import pandas as pd
import shutil
import stat

#File location for BOMs
XLS_source = "C:\LAC Conveyors\Playground\Playground Kon\Design Table Trial\\"
#File location for Drawings
DRW_source = r"C:\LAC Conveyors\Playground\Playground Kon\Design Table Trial\Drawing Detail\\"
#File location for library
End_source = r"C:\LAC Conveyors\Playground\Playground Chris Yates\Python Tests\\"

def CopyFiles(source_folder, destination_folder, BOM_folder):
    prtName = FindExcelDocs(BOM_folder)
    num = 0
    while num < len(File_name):
        for file in os.listdir(source_folder):
            source = source_folder + file
            for BOM in prtName:
                dest = destination_folder[:-2] + "\\" + File_name[num] + "\\"
                destination = dest + file
                for item in BOM:
                    draw = str(item) + '.SLDDRW'
                    if os.path.isfile(source) and draw == file:
                        try:
                            shutil.copy(source, destination)
                            print("Copied ", file)
                            continue
                        except:
                            os.chmod(destination, stat.S_IWRITE)
                            shutil.copy(source, destination)
                            print("Updated ", file)
        num += 1

def FindExcelDocs(source):
    global File_name
    BOMs = []
    File_name = []
    for file in os.listdir(source):
        if file.endswith(".xls") or file.endswith(".xlsx"):
            CreateFolder(file, End_source)
            File_name.append(dirName)
            df = pd.read_excel(source + file, usecols='B')
            BOM = []
            for num in range(len(df)):
                BOM.append(df.iloc[num][0])
            BOMs.append(BOM)
    return(BOMs)
                    
def CreateFolder(name, location):   
    global dirName
    try:
        dirName = name.strip('.xlsx')
    except:
        dirName = name.strip('.xls')
        
    path = os.path.join(location, dirName)
    
    try:
        os.mkdir(path)
        print("Folder '% s' created" % dirName)
    except OSError as error:
        pass
        
CopyFiles(DRW_source, End_source, XLS_source)
