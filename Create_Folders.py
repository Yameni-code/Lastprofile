import os
"""
create_Folder: creates a folder at giving path
parameter1: path to were the folder will be created
parameter2: folder name
retreval:  
"""
def create_Folder(parent_path, folder_name):
    path = os.path.join(parent_path, folder_name)
    os.mkdir(path, 0o666)
    print("Folder '% s' created" % folder_name)

def run_fun():
    for kw in range(5, 52):
        parent_path = "//cifs02/RoamingData$/u2110370/Documents/Standardlastprofilen"
        folder_name = "kw_" + str(kw)
        create_Folder(parent_path, folder_name)