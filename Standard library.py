# Import the below libraries to use their methods globally in any function
import os
import pandas as pd
import shutil
import stat
import string

# List containing all uppercase alphabet letters
alphabet = list(string.ascii_uppercase)

# File location for macro generated drawings
DRW_source = r"C:\LAC Conveyors\Standard Module CAD Library\Pallet\Roller Straight\Sections\Library Generation Related\Macro-Generated Drawings\\"
# File location for issued drawings
End_source = r"C:\LAC Conveyors\Standard Module CAD Library\Pallet\Roller Straight\Sections\Issued Standard Drawings\\"

#------------------------------------------------------------ 
# Primary function that runs all sub functions
def CopyFiles(source_folder, destination_folder):
    
    # List containing names of each BOM, also creates folder for each
    BOM_List = FindExcelDocs(source_folder)
    
    # For each file in the generated folder obtain its location name
    for file in os.listdir(source_folder):
        source = source_folder + file
        
        # For every BOM, set the destination folder to the corresponding name
        for BOM in BOM_List:
            num = BOM_List.index(BOM)
            dest = destination_folder[:-2] + "\\" + File_name[num] + "\\"
            # Self explanatory, see function for more info
            CreateSubFolders(dest)
            
            # For each part in the BOM copy it to a folder based on the identifier (e.g -L)
            for item in BOM:
                sort(item, source, dest, file, 0) # Copy lasered
                sort(item, source, dest, file, 1) # Copy Machined
                sort(item, source, dest, file, 2) # Copy Fabrications
                sort(item, source, dest, file, 3) # Copy Rollers
                sort(item, source, dest, file, 4) # Copy Misc
                sort(item, source, dest, file, 5) # Copy Plastics
                sort(item, source, dest, file, 6) # Copy Sub-Assembly
                
    # Copy the top-level assembly drawing at the end
    CopyAssy(source_folder, destination_folder)
    CopyBOM(source_folder, destination_folder)
    print("Operation completed")
#------------------------------------------------------------ 
def FindExcelDocs(source):
    
    # Sets the File_name list to global so it can be used in other functions
    global File_name
    
    BOMs = []
    File_name = []
    
    # If any file is an excel sheet (BOM), create a folder for it and add its name to the File_name list
    for file in os.listdir(source):
        if file.endswith(".xls") or file.endswith(".xlsx"):
            CreateFolder(file, End_source)
            File_name.append(dirName)
            df = pd.read_excel(source + file, usecols="B")
            BOM = []
            
            # Removes indent spaces from each entry of the BOM
            df = df.replace(
                to_replace=' ', 
                value='',
                regex=True)
            
            # Add each part from a BOM to its own list then add this full list to the giant BOMs list
            for num in range(len(df)):
                BOM.append(df.iloc[num][0])
            BOMs.append(BOM)
            
    # The function finally outputs a list of BOMs containing all parts from each individual BOM      
    return(BOMs)
#------------------------------------------------------------                    
def CreateFolder(name, location):   
    
    # Allows dirName to be used outside of the function (see above)
    global dirName
    
    # Remove the excel extension from a BOM (leaving just its name)
    try:
        dirName = name.strip(".xlsx")
    except:
        dirName = name.strip(".xls")
    
    # Location is used when the function is called and therefore is defined elsewhere
    # Creates the destination (path) for creating the folder
    path = os.path.join(location, dirName)
    
    # Attempt to create the folder if this is the first time or skip as it cant be made again
    try:
        os.mkdir(path)
        print("Folder '% s' created" % dirName)
    except OSError as error:
        pass
#------------------------------------------------------------         
def CopyAssy(source_folder, destination_folder):
    
    # Sets the index position in the list [File_name] for each entry (a.k.a Assembly)
    for File in File_name:
        i = File_name.index(File)
        
        # Cycle each letter in the alphabet starting at [A] then..
        # Cycle each file in the generated folder and set its current location & destination..
        # Attempt to copy the assembly if the drawing for it existis as a .pdf
        for letter in alphabet:
            for file in os.listdir(source_folder):            
                source = source_folder + file
                dest = destination_folder[:-2] + "\\" + File_name[i] + "\\"
                destination = dest + file
                # Attempt to create a valid assembly name
                Assy = str(File_name[i]) + "-" + str(letter) + ".pdf"
                
                # If the name we created matches a real assembly it will copy or update existing
                if Assy == file:
                    try:
                        os.chmod(destination, stat.S_IWRITE)
                        shutil.copy(source, destination)
                        print("Updated ", file)
                    except:
                        shutil.copy(source, destination)
                        print("Copied ", file)
#------------------------------------------------------------         
def CreateSubFolders(location):
    
    # List of sub-folders we need to generate
    Type = ["Fabrications", "Lasered", "Machined", "Rollers", "Misc", "Plastic", "Sub-Assembly"]
    
    # Uses location defined elsewhere and attaches the sub-folder extension to create the folder
    for i in Type:
        path = os.path.join(location, i)   
        try:
            os.mkdir(path)
            print("Folder '% s' created" % i)
        except OSError as error:
            pass
#------------------------------------------------------------ 
def sort(item, source, dest, file, option):
    
    # List of possible file types (PDF included twice due to some files using uppercase)
    Types = [".pdf", ".dxf", ".DXF", ".PDF"]
    
    # Lists containing sub-folders and identifier types - must be corresponding
    for Type in Types:
        i = ["-L", "-M", "-F", "-R", "-D", "-P", "-A"]
        Folders = ["Lasered\\", "Machined\\", "Fabrications\\", "Rollers\\", "Misc\\", "Plastic\\", "Sub-Assembly\\"]
        
        # Gets the latest revision if the file is a PDF (not necessary for DXFs etc)
        if Type == ".pdf" or ".PDF":
            for letter in alphabet:
                j = alphabet.index(letter)
                # Attempt to create a valid file name
                doc = str(item) + "-" + alphabet[j] + Type
                
                # If our created name matches a real file, copy it to the folder corresponding
                # with its identifier letters at the end
                if doc == file:
                    if item[-2:] == i[option]:
                        destination = dest + Folders[option] + file
                        try:
                            os.chmod(destination, stat.S_IWRITE)
                            shutil.copy(source, destination)
                            print("Updated ", file)
                        except:
                            shutil.copy(source, destination)
                            print("Copied ", file)
                            
        # No need to cycle revisions so repeat the above without a rev at the end (e.g "-A")
            else:
                doc = str(item) + Type
                if doc == file:
                    if item[-2:] == i[option]:
                        destination = dest + Folders[option] + file
                        try:
                            os.chmod(destination, stat.S_IWRITE)
                            shutil.copy(source, destination)
                            print("Updated ", file)
                        except:
                            shutil.copy(source, destination)
                            print("Copied ", file)
#------------------------------------------------------------
def CopyBOM(source_folder, destination_folder):
        # Sets the index position in the list [File_name] for each entry (a.k.a BOM)
    for File in File_name:
        i = File_name.index(File)

        # Cycle each file in the generated folder and set its current location & destination..
        # Attempt to copy the BOM if the file for it existis as a .xls or .xlsx
        for file in os.listdir(source_folder):            
            source = source_folder + file
            dest = destination_folder[:-2] + "\\" + File_name[i] + "\\"
            destination = dest + file
            # Attempt to create a valid BOM name
            BOM_1 = str(File_name[i]) + ".xls"
            BOM_2 = str(File_name[i]) + ".xlsx"

            # If the name we created matches a real BOM it will copy or update existing
            if BOM_1 == file or BOM_2 == file:
                try:
                    os.chmod(destination, stat.S_IWRITE)
                    shutil.copy(source, destination)
                    print("Updated ", file)
                except:
                    shutil.copy(source, destination)
                    print("Copied ", file)  
#------------------------------------------------------------

# EXECUTE ORDER 66...
CopyFiles(DRW_source, End_source)
