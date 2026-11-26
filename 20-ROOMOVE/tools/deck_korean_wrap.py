# 한글 단어가 중간에서 잘리지 않도록 모든 문단에 eaLnBrk="0"
import sys,zipfile,shutil,re
for f in sys.argv[1:]:
    zin=zipfile.ZipFile(f); zout=zipfile.ZipFile(f+'.tmp','w',zipfile.ZIP_DEFLATED)
    for i in zin.infolist():
        d=zin.read(i.filename)
        if re.match(r'ppt/(slides/slide|slideLayouts/slideLayout|slideMasters/slideMaster)\d+\.xml$',i.filename):
            x=d.decode('utf-8')
            x=re.sub(r'<a:pPr(?![^>]*eaLnBrk)',r'<a:pPr eaLnBrk="0"',x)
            x=re.sub(r'<a:p>(?!<a:pPr)',r'<a:p><a:pPr eaLnBrk="0"/>',x)
            x=x.replace('lang="en-US"','lang="ko-KR"')
            d=x.encode('utf-8')
        zout.writestr(i,d)
    zin.close(); zout.close(); shutil.move(f+'.tmp',f)
