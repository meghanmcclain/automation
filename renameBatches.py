"""renamed ~13,576 files in different folders and subfolders and put them all in one folder"""

import os
import shutil

FOLDER_PATH = r'K:/Batches To_film'
count = 0

for root,directories,files in os.walk(FOLDER_PATH,topdown=False):
    for fileName in files:
        if fileName.endswith(".tif"):
            #fullPath = path 
            full_path_with_drive_and_slashes = os.path.join(root,fileName)
            print(root)
            print(count)
            name_without_slashes = full_path_with_drive_and_slashes.replace("\\","")
            fullPath = name_without_slashes.replace("K:/","")
            print(fullPath)
            print(full_path_with_drive_and_slashes)
            print(full_path_with_drive_and_slashes.replace(fileName,fullPath))
            shutil.copyfile(full_path_with_drive_and_slashes,r'K:/RenamedBatches/'+fullPath.replace("Batches To_film", ""))
            print("Tif file "+fileName+" renamed to "+fullPath)
            count += 1
            continue
        else:
            continue
