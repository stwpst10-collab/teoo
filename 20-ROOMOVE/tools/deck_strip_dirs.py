import sys,zipfile,shutil
for f in sys.argv[1:]:
    zin=zipfile.ZipFile(f); zout=zipfile.ZipFile(f+'.tmp','w',zipfile.ZIP_DEFLATED)
    for i in zin.infolist():
        if not i.filename.endswith('/'): zout.writestr(i, zin.read(i.filename))
    zin.close(); zout.close(); shutil.move(f+'.tmp',f)
