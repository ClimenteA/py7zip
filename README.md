# py7zip
Unzipping files with 7zip and python fast!

The main idea of this module is this:


```

import subprocess

def execute7z(zipPathInput, zipPassword, zipPathOutput):
    """Excecute 7zip batch command using subprocess call"""
    process = subprocess.Popen([r"7z/7za.exe", "x", zipPathInput, "-p{}".format(zipPassword), "-o{}".format(zipPathOutput)])
    process.wait()
    process.kill()
    return 0
    
    
#do stuff
#do stuff
#do stuff

process = execute7z(zfp, password, unzip_files_paths)

if process == 0:
    #unzipping done
    #do some checks
    #etc
   
```

You can add a new arg to execute7z where you can put diferent 7zip CLI commands.<br/>
Go here https://www.7-zip.org/download.html and download the standalone console version.<br/>
Go here https://www.dotnetperls.com/7-zip-examples for more 7zip commands.<br/>
