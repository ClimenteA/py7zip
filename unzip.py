import re, os, sys
import subprocess
from zipfile import ZipFile
import pandas as pd



print("\n\nMake sure you have zipPassword.xlsx next to the program and you have sufficient space on disk\n")


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
    process = subprocess.Popen([r"7z/7za.exe", "x", zipPathInput, "-p{}".format(zipPassword), "-o{}".format(zipPathOutput)])
    process.wait()
    process.kill()
    return 0

                                    

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




def getDirSize(dirPath):
    """Get a directory size"""
    total_size = 0
    for path, dirs, files in os.walk(dirPath):
        for f in files:
            fp = os.path.join(path, f)
            total_size += os.path.getsize(fp)
    return int(total_size)

def getfileSize(filepath):
    """Get file size"""
    metadata = os.stat(filepath)
    file_size = int(metadata.st_size) # bytes
    
    return file_size

def getSize(apath):
    """Get file or dir size"""
    if os.path.isfile(apath):
        return getfileSize(apath)
    elif os.path.isdir(apath):
        return getDirSize(apath)
    else:
        raise ValueError("Input is neither a path or a file!")


def checkZIP(zipFilePath, unzippedFilePath):
    """Check if zip extracted corectly by verifying zip and unzipped size"""
    zip_size = getSize(zipFilePath)
    unzip_size = getSize(unzippedFilePath)
     
    #print("zip_size = {}, unzip_size = {}".format(zip_size, unzip_size))
    
    if unzip_size < zip_size/2:
        return False
    else:
        return True
        


try:
    print("\n\nPlease wait...\n")
    createUnzipDirs(unzip_files_paths)
    
    try:
        os.remove("errors.txt")
    except:
        pass

    err = open("errors.txt", "a")


    passdf = pd.read_excel("zipPassword.xlsx")
    passdf_idx = passdf.index.tolist()

    nokli = []
    for j, zfp in enumerate(zip_files_paths):
        zfName = zfp.split("\\")[-1]
        nestedYear_found = nestedYear(zfp)
        
        for i in passdf_idx:
            try:
                ac = passdf.loc[i, "A/C"]
                year = passdf.loc[i, "Year"]
                password = passdf.loc[i, "Password"]

                if re.search(str(ac), zfName) and str(year) == str(nestedYear_found):
                    process = execute7z(zfp, password, unzip_files_paths[j])
                    if process == 0:
                        if checkZIP(zfp, unzip_files_paths[j]):
                            print("\nCool unzipping it's ok!\n")
                            msg = "Unzipping ok for > {}\n".format(zfp)
                            err.write(msg)
                            #idxli.append(i)
                        else:
                            print("\nPlease check zip > ", zfp)
                            errMsg = str("\nWarning: Got a big diference in sizes!\nPlease check zip > {} \nif it was unzipped corectly in > {}\n\n".format(zfp, unzip_files_paths[j]))
                            err.write(errMsg)
                            nokli.append(str("\nPlease check: " + str(zfp) + ' and ' + str(unzip_files_paths[j])))
            except:
                try:#not tested
                    zipName = passdf.loc[i, "Name"]
                    password = passdf.loc[i, "Password"]
                except:
                    zipName = zfName
                    password = ''

                if zipName == zfName:
                    process = execute7z(zfp, password, unzip_files_paths[j])
                    if process == 0:
                        if checkZIP(zfp, unzip_files_paths[j]):
                            print("\nCool unzipping it's ok!\n")
                            msg = "\nUnzipping ok for > {}\n".format(zfp)
                            err.write(msg)
                            #idxli.append(i)
                        else:
                            print("\nPlease check zip > ", zfp)
                            errMsg = str("\nWarning: Got a big diference in sizes!\nPlease check zip > {} \nif it was unzipped corectly in > {}\n\n".format(zfp, unzip_files_paths[j]))
                            err.write(errMsg)
                            nokli.append(str("\nPlease check: " + str(zfp) + ' and ' + str(unzip_files_paths[j])))


    for n in nokli:
        print(n)

    err.close()
    print("\n\nDone!\nFor any issue contact AtexisRO: alin.climente@atexis.eu")
    input("Press enter to exit..")
except Exception as e:
    err.close()
    print("Got this error: ", e)
    input("Press enter to exit..")





