# Import the below libraries to use their methods globally in any function
import os
import pandas as pd
import shutil
import stat
import string

# List containing all uppercase alphabet letters starting from Z
alphabet = list(string.ascii_uppercase)
alphabet = alphabet[::-1]

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
        print("Working on " + file)
        print("")
        source = source_folder + file
        
        # For every BOM, set the destination folder to the corresponding name
        for BOM in BOM_List:
            num = BOM_List.index(BOM)
            dest = destination_folder[:-2] + "\\" + File_name[num] + "\\"
            # Self explanatory, see function for more info
            CreateSubFolders(dest)
            
            # For each part model in the BOM copy it to a folder based on the identifier (e.g -L)
            for item in BOM:
                sort(item, source, dest, file, 0) # Copy lasered
                sort(item, source, dest, file, 1) # Copy Machined
                sort(item, source, dest, file, 2) # Copy Fabrications
                sort(item, source, dest, file, 3) # Copy Rollers
                sort(item, source, dest, file, 4) # Copy Misc
                sort(item, source, dest, file, 5) # Copy Plastics
                sort(item, source, dest, file, 6) # Copy Sub-Assembly

    # Copy the top-level assembly drawing and BOM at the end
    CopyAssy(source_folder, destination_folder)
    CopyBOM(source_folder, destination_folder)
    print("")
    # Delete any old Revs
    RevControl(destination_folder)
    print("")
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
            
    # The function finally outputs a list of BOMs, containing all parts from each individual BOM      
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
    
    # Attempt to create the folder if this is the first time or skip as it cant be made if it exists
    try:
        os.mkdir(path)
        print("Folder '% s' created" % dirName)
    except OSError as error:
        pass
#------------------------------------------------------------         
def CopyAssy(source_folder, destination_folder):
    
    # Sets the index position in the list [File_name] for each entry (a.k.a Assembly)
    # File_name is global now from a previous function which is how we access it here
    for File in File_name:
        x = []
        i = File_name.index(File)
        
        # Cycle each letter in the alphabet starting at [Z] then..
        # Cycle each file in the generated folder and set its current location & destination..
        # Attempt to copy the assembly if the drawing for it existis as a .pdf
        for letter in alphabet:
            for file in os.listdir(source_folder):            
                source = source_folder + file
                dest = destination_folder[:-2] + "\\" + File_name[i] + "\\"
                destination = dest + file
                # Attempt to create a valid assembly name
                Assy = str(File_name[i]) + "-" + str(letter) + ".pdf"

                # If the name we created matches a real assembly it will copy latest if it isn't already there
                # Also adds each rev to a list called [x]
                if Assy == file:
                    x.append(Assy)
                    try:
                        shutil.copy(source, destination)
                        print("copied ", file)
                        break
                    except:
                        ("Latest Rev of " + file + " exists")
                        pass
        
        # Still working per file in the list, if there are more than one rev of this file..
        if len(x) > 1:
            # Excluding latest rev by strating at 1, cycle the old revs
            # For each old rev set its location and remove it
            for old_rev in range(1, len(x)):
                location = destination_folder[:-1] + File + "\\" + x[old_rev]
                try:
                    os.chmod(location, stat.S_IWRITE)
                    os.remove(location)
                    print("Removed " + x[old_rev])
                except:
                    continue
#------------------------------------------------------------         
def CreateSubFolders(location):
    
    # List of sub-folders we need to generate
    Type = ["Fabrications", "Lasered", "Machined", "Rollers", "Misc", "Plastic", "Sub-Assembly"]
    
    # Uses location defined earlier and attaches the sub-folder extension to create the folder
    for i in Type:
        path = os.path.join(location, i)
        try:
            os.mkdir(path)
            print("Folder '% s' created" % i)
        except OSError as error:
            pass
#------------------------------------------------------------ 
def sort(item, source, dest, file, option):
    
    # List of possible file types (PDF & DXF included twice due to some files using uppercase)
    Types = [".pdf", ".dxf", ".DXF", ".PDF"]
    
    # Lists containing sub-folders and identifier types - MUST be corresponding
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
                # with its identifier letters at the end (i.e lasered for -L)
                if doc == file:
                    if item[-2:] == i[option]:
                        destination = dest + Folders[option] + file
                        try:
                            shutil.copy(source, destination)
                            print("Copied ", file)
                            break
                        except:
                            print("Latest Rev of " + file + " exists")
                            break
                  
        # No need to cycle revisions on DXF so repeat the above without a rev at the end (e.g "-A")
        # Updating the file to the latest or copying if one doesn't exist
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

            # If the name we created matches a real BOM it will copy to folder or update existing
            if BOM_1 == file or BOM_2 == file:
                try:
                    os.chmod(destination, stat.S_IWRITE)
                    shutil.copy(source, destination)
                    print("Updated ", file)
                except:
                    shutil.copy(source, destination)
                    print("Copied ", file)  
#------------------------------------------------------------
def RevControl(destination_folder):
    
    # Add each folder in the issued directory to the list [x]
    x = []
    for Dir in os.listdir(destination_folder):
        x.append(Dir + "\\")
    
    # For each of these fill up the list [y] with the sub folders (e.g misc, lasered etc)
    for Folder in x:
        y = []
        dest = destination_folder + Folder
        for sub_folder in os.listdir(dest):
            if sub_folder[-3:] != "pdf" and sub_folder[-3:] != "xls":
                y.append(sub_folder + "\\")
        
        # Within each subfolder if it contains PDFs add them to the list [z]
        for sub_folder in y:
            z = []
            for file in os.listdir(destination_folder + Folder + sub_folder):
                if file[-3:] == "pdf":
                    z.append(file) 
            
            # If [z] contains files cycle through each and assign its rev to 'R'
            if z != []:
                Bin = []
                # Take the current rev and add the previous rev to the [Bin] list if possible
                # Wont work on Rev-A, will skip this file for Rev-A
                for instance in z:
                    R = alphabet.index(instance[-5])
                    Rev = ""
                    try:
                        R += 1
                        Rev = alphabet[R]
                    except:
                        continue
                    finally:
                        old_rev = instance[:-6] + "-" + Rev + ".pdf"
                        Bin.append(old_rev)
                
                # For each old rev in the [Bin] attempt to remove it
                for PDF in range(len(Bin)):
                    location = destination_folder[:-1] + Folder + sub_folder + Bin[PDF]
                    try:
                        os.chmod(location, stat.S_IWRITE)
                        os.remove(location)
                        print("Removed " + Bin[PDF])
                    except:
                        continue
#------------------------------------------------------------

# EXECUTE ORDER 66...
CopyFiles(DRW_source, End_source)
