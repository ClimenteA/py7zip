
# coding: utf-8

# In[1]:


import re, os
import subprocess
from zipfile import ZipFile
import pandas as pd




print("\n\nMake sure you have zipPassword.xlsx next to the program\n")


zippedPath = input("\n\nInsert the path to 'zipped files' below:\n")

#zippedPath = r"E:\_python_macro\Unzip zip with pass\zipped files"
unzippedPath = zippedPath.replace("zipped files", "unzipped files")


months = ['JANUARY',
          'FEBRUARY',
          'MARCH',
          'APRIL',
          'MAY',
          'JUNE',
          'JULY',
          'AUGUST',
          'SEPTEMBER',
          'OCTOBER',
          'NOVEMBER',
          'DECEMBER']

files = os.listdir(zippedPath)
zip_files = [z for z in files if re.search(".zip", z)] 
zip_files_paths = [os.path.join(zippedPath, zp) for zp in zip_files]
unzip_files_paths = [z.replace("zipped files", "unzipped files").replace(".zip", "").replace("Downloaded_documents_", "") for z in zip_files_paths]


def createUnzipDirs(unzip_files_paths):
    os.mkdir(unzippedPath)
    for dirName in unzip_files_paths:
        os.mkdir(dirName)



def execute7z(zipPathInput, zipPassword, zipPathOutput):
    """Excecute 7zip batch command using subprocess call
       x {} - full path to zip; -p"{}" - password; -o{} - output path of unzip
    """#'7z x {} -p"{}" -o{}'
    subprocess.Popen([r"7z/7za.exe", "x", zipPathInput, "-p{}".format(zipPassword), "-o{}".format(zipPathOutput)])


def zipFilesNames(zipName):  
    """Get a list of names inside that zip files"""
    file = ZipFile(zipName)
    filenames = []
    for name in file.namelist():
        filenames.append(name)
    return filenames


def getYear(zipName):
    """Search for a month year pattern if found return year"""
    for month in months:
        pattern = month + " " +"\\d{4}"
        match = re.search(pattern, zipName)
        if match:
            year = str(match.group().split(' ')[-1].strip())
            return year


def nestedYear(zipFilePath):
    """Use zipFilesNames and getYear to return the year from the nested zip file"""
    filesName = zipFilesNames(zipFilePath)
    for fn in filesName:
        year = getYear(fn)
        if isinstance(year, str):
            return year


try:
    print("\n\nPlease wait...\n")
    createUnzipDirs(unzip_files_paths)

    passdf = pd.read_excel("zipPassword.xlsx")
    passdf_idx = passdf.index.tolist()

    for i in passdf_idx:
        ac = passdf.loc[i, "A/C"]
        year = passdf.loc[i, "Year"]
        password = passdf.loc[i, "Password"]

        for j, zfp in enumerate(zip_files_paths):
            zfName = zfp.split("\\")[-1]
            nestedYear_found = nestedYear(zfp)
            if re.search(str(ac), zfName) and str(year) == str(nestedYear_found):
                execute7z(zfp, password, unzip_files_paths[j])
    
    #print("\n\nDone!\n For any issue contact AtexisRO: alin.climente@atexis.eu")
    #input("Press enter to exit..")
    input()
except Exception as e:
    #print("Got this error: ", e)
    #input("Press enter to exit..")
    input()
