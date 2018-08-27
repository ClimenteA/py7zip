import re, os, sys
import subprocess
import shutil
from zipfile import ZipFile
import pandas as pd
import numpy as np



def get_paths(zippedPath):
    """Construct paths to zipped files and unzipped files"""
    
    unzippedPath = zippedPath.replace("zipped files", "unzipped files")

    files = os.listdir(zippedPath)
    zip_files = [z for z in files if re.search(".zip", z)] 

    if len(files) != len(zip_files):
        print("\nWarning you have files that are not in .zip format!")

    zip_files_paths = [os.path.join(zippedPath, zp) for zp in zip_files]
    unzip_files_paths = [z.replace("zipped files", "unzipped files").replace(".zip", "") for z in zip_files_paths]
    
    return unzippedPath, unzip_files_paths, zip_files_paths


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


def zipFilesNames(zipName):  
    """Get a list of names inside that zip files"""
    file = ZipFile(zipName)
    filenames = []
    for name in file.namelist():
        filenames.append(name)
    file.close()
    
    return filenames


def getZipInfo(zip_files_paths):
    """Get from folder that contains zipped files the size and dir tree for each zip """
    zipInfo = {}
    for zfp in zip_files_paths:
        #print("\nGet size/tree from:", zfp)
        zipName = zfp.split('\\')[-1]
        info = {'path': zfp, 'size': getSize(zfp), 'tree':zipFilesNames(zfp)}
        zipInfo[zipName] = info
    return zipInfo


def dfZipInfo(zip_files_paths):    
    """Create a df with the information from the zipInfo dict"""
    
    zipInfo = getZipInfo(zip_files_paths)

    zname = []
    zsize = []
    ztree = []
    zpath = []
    for zipName, zipDict in zipInfo.items():
        size = zipInfo[zipName]['size']
        tree = zipInfo[zipName]['tree']
        path = zipInfo[zipName]['path']

        zname.append(zipName)
        zsize.append(size)
        ztree.append(" ".join(sorted(tree)))
        zpath.append(path)

    df = pd.DataFrame({})
    df['zipName'] = pd.Series(zname)
    df['zipPass'] = np.nan
    df['zipPath'] = pd.Series(zpath)
    df['zipSize'] = pd.Series(zsize)
    df['zipTree'] = pd.Series(ztree)
    df['zipDups'] = pd.Series(df.index.tolist())
    
    return df


def solve_zip_duplicates(zip_files_paths):
    """Filter from the df the zip files that are duplicates"""
    
    zipInfodf = dfZipInfo(zip_files_paths)

    sizeGroup = zipInfodf.groupby(['zipSize'])
    treeGroup = zipInfodf.groupby(['zipTree'])

    for size, sdf in sizeGroup:
        sdf_idxli = sdf.index.tolist()
        if len(sdf_idxli) > 1:
            #print(sdf_idxli)
            for i in sdf_idxli:
                zipInfodf.loc[i, 'zipDups'] = 'dups {}'.format(str(sdf_idxli))
                zName = zipInfodf.loc[i, 'zipName']

    for tree, tdf in treeGroup:
        tdf_idxli = tdf.index.tolist()
        if len(tdf_idxli) > 1:
            #print(tdf_idxli)
            for i in tdf_idxli:
                zipInfodf.loc[i, 'zipDups'] = 'dups {}'.format(str(tdf_idxli))
                zName = zipInfodf.loc[i, 'zipName']


    zipInfodf.drop_duplicates(subset='zipDups', keep='last', inplace=True)
    zipInfodf.reset_index(drop=True, inplace=True)

    return zipInfodf


def get_passwords():
    """Open and get passwords in the excel next to executable. Make sure that is has column Password"""
    cwd = os.getcwd()
    files = os.listdir(cwd)
    xlfile = [f for f in files if f == 'passwords.xlsx']
    if len(xlfile) == 1:
        passdf = pd.read_excel(xlfile[0])
        if 'Passwords' not in passdf.columns.tolist():
            raise Exception("The .xlsx file must have a column named Pasword")
        return passdf['Passwords'].tolist()
    else:
        raise Exception("No '.xlsx' file found next to the executable!")


def deletetree(apath):
    """Delete folder and all of it's contents"""
    try:
        shutil.rmtree(apath) 
        return True
    except:
        shutil.rmtree(apath, ignore_errors=True) #delete files that are not opened
        response = "Could not remove folder {}, because some files are opened! Please delete them manually...".format(apath)
    return response

def createDir_bin():
    """Create bin folder next to executable"""
    try:
        os.mkdir("bin")
    except:
        pass




def checkPassforZip(zip_file_path, passwordsli, re_search_criteria=".pdf"):
    """Check for a zip file what password works by extracting one file"""
    zip_file_path = r"{}".format(zip_file_path)
    files = zipFilesNames(zip_file_path)
    file = [f for f in files if re.search(re_search_criteria, f)]
    if len(file) != 0:
        file = file[0]
    
    #zip_file_path = r"E:\_python_macro\Unzip zip with pass\zipped files\Downloaded_documents_B-HLA_17_08_2018_12_12_00.zip"
    #p = "F367718:33:28"

    for password in passwordsli:
        pstr = "{}".format(str(password).strip())
        createDir_bin()
        try:
            
            with ZipFile(zip_file_path, 'r') as z:
                z.extract(file, path="bin", pwd=bytes(pstr.encode()))
                
            deletetree('bin')
            return password.strip()
        
        except Exception as e:
            
            if re.search('Bad password for file', str(e)):
                #print("Bad password for zip file!")
                deletetree('bin')
                z.close()
            else:
                errmsg = "{}".format(str(e))
                deletetree('bin')
                z.close()
                raise Exception(errmsg)



def fillPasswords(zipdf):
    """Fill df with passwords for each zip if found"""
    
    passwordsli = get_passwords()
    
    zfp_with_errors = []
    for i, zfp in enumerate(zipdf['zipPath'].tolist()):
        print("\nSearching password for: ", zfp)
        try:
            password = checkPassforZip(zfp, passwordsli)
            if isinstance(password, str):
                zipdf.loc[i, 'zipPass'] = password
            else:
                zipdf.loc[i, 'zipPass'] = 'Password not found'
                zfp_with_errors.append(zfp)
        except Exception as e:
            print("Got: {} for: {}\n".format(str(e), zfp))
            zipdf.loc[i, 'zipPass'] = 'Password not found'
            zfp_with_errors.append(zfp)

    print("\n\nCan't find passwords for these files:")
    for zp in zfp_with_errors:
        print('\n',zp)

    zipdf.to_excel("zipInfo.xlsx", index=False)


def checkZIP(zipFilePath, unzippedFilePath):
    """Check if zip extracted corectly by verifying zip and unzipped size"""
    zip_size = getSize(zipFilePath)
    unzip_size = getSize(unzippedFilePath)
     
    #print("zip_size = {}, unzip_size = {}".format(zip_size, unzip_size))
    
    if unzip_size < zip_size/2:
        return False
    else:
        return True


def execute7z(zipPathInput, zipPassword, zipPathOutput):
    """Excecute 7zip batch command using subprocess call
       x {} - full path to zip; -p"{}" - password; -o{} - output path of unzip
    """#'7z x {} -p"{}" -o{}'
    process = subprocess.Popen([r"7z/7za.exe", "x", zipPathInput, "-p{}".format(zipPassword), "-o{}".format(zipPathOutput)])
    process.wait()
    process.kill()
    return 0



def masterUnzip(unzippedPath):
    zipdf = pd.read_excel("zipInfo.xlsx")
    zipdf = zipdf[zipdf["zipPass"] != 'Password not found']
    zipdf.reset_index(drop=True, inplace=True)

    try:
        os.remove("errors.txt")
    except:
        pass

    err = open("errors.txt", "a")

    try:
        os.mkdir(unzippedPath)
    except:
        print("Please delete or move 'unzipped files' folder!")
        input("Press enter to exit...")
        sys.exit()

    for i, zfp in enumerate(zipdf["zipPath"].tolist()):
        unzip_file_path = zfp.replace("zipped files", "unzipped files")
        password = zipdf.loc[i, "zipPass"]

        os.mkdir(unzip_file_path)

        process = execute7z(zfp, password, unzip_file_path)

        if process == 0:
            if checkZIP(zfp, unzip_file_path):
                print("\nCool unzipping it's ok!\n")
                msg = "Unzipping ok for > {}\n".format(zfp)
                err.write(msg)
                #idxli.append(i)
            else:
                print("\nPlease check zip > ", zfp)
                errMsg = str("\nWarning: Got a big diference in sizes!\nPlease check zip > {} \nif it was unzipped corectly in > {}\n\n".format(zfp, unzip_file_path))
                err.write(errMsg)

    err.close()



##################################
############# START #############
##################################



print("\n\nMake sure you have zipPassword.xlsx next to the program and you have sufficient space on disk\n")


zippedPath = input("\n\nInsert the path to 'zipped files' below:\n")
#zippedPath = r"E:\_python_macro\Unzip zip with pass\zipped files"
try:
    unzippedPath, unzip_files_paths, zip_files_paths = get_paths(zippedPath)
except Exception as e:
    print("Path inserted can't be processed! Got: ", e)
    input("Press enter to exit...")
    sys.exit()

skip = input("\nMake sure you have 'zipInfo.xlsx' filed corectly if you skip!\nSkip checks for duplicates and passwords checks?(y/n)\n")

if skip.lower() == "n" or skip.lower() == "":
    print("\nPlease wait..")
    zipdf = solve_zip_duplicates(zip_files_paths)
    fillPasswords(zipdf)
    try:
        masterUnzip(unzippedPath)
        print("\n\nDone!\nFor any issue contact AtexisRO: alin.climente@atexis.eu")
        input("Press enter to exit..")
    except Exception as e:
        print("\n\nGot this error: ", e)
        input("Press enter to exit...")
elif skip.lower() == "y":
    print("\nPlease wait..")
    try:
        masterUnzip(unzippedPath)
        print("\n\nDone!\nFor any issue contact AtexisRO: alin.climente@atexis.eu")
        input("Press enter to exit..")
    except Exception as e:
        print("\n\nGot this error: ", e)
        input("Press enter to exit...")
else:
    print("You can insert only y - yes and n - for no ")
