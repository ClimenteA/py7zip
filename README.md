# py7zip
Unzipping files with 7zip and python fast!

The modules is for a custom use the main func of the module is this one:
```

def execute7z(zipPathInput, zipPassword, zipPathOutput):
    """Excecute 7zip batch command using subprocess call
       x {} - full path to zip; -p"{}" - password; -o{} - output path of unzip
    """#'7z x {} -p"{}" -o{}'
    process = subprocess.Popen([r"7z/7za.exe", "x", zipPathInput, "-p{}".format(zipPassword), "-o{}".format(zipPathOutput)])
    process.wait()
    process.kill()
    return 0
    
    
#do stuff
#do stuff
#do stuff

process = execute7z(zfp, password, unzip_files_paths[j])

if process == 0:
    #unzipping done
    #do some checks
    #etc
   
```
