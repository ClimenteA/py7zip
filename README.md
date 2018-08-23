# py7zip
Unzipping files with 7zip and python fast!

The modules is for a custom use the main func of the module is this one:

def execute7z(zipPathInput, zipPassword, zipPathOutput):
    """Excecute 7zip batch command using subprocess call
       x {} - full path to zip; -p"{}" - password; -o{} - output path of unzip
    """#'7z x {} -p"{}" -o{}'
    subprocess.Popen([r"7z/7za.exe", "x", zipPathInput, "-p{}".format(zipPassword), "-o{}".format(zipPathOutput)])
