import os
import sys
import zipfile


def functionDecode(fromCode, toCode):
    reload(sys)
    sys.setdefaultencoding(fromCode)
    print sys.getdefaultencoding()

    print "Processing File " + sys.argv[1]

    file = zipfile.ZipFile(sys.argv[1], "r");

    for name in file.namelist():
        utf8name = name.decode(toCode)
        print "Extracting " + utf8name
        pathname = os.path.dirname(utf8name)
        if not os.path.exists(pathname) and pathname != "":
            os.makedirs(pathname)
        data = file.read(name)
        if not os.path.exists(utf8name):
            fo = open(utf8name, "w")
            fo.write(data)
            fo.close
    file.close

# functionDecode("utf-8", "utf-8")


fromCodeList = ["utf-8","big5","big5hkscs","cp037","cp424","utf_8_sig","gbk","gb18030","hz","cp950","iso2022_jp_2","utf_32","utf_16"]

for fromCodea in fromCodeList:
    for toCodea in fromCodeList:
        try:
            functionDecode(fromCodea, toCodea)
        except:
            print "error"
        finally:
            print "always excute"
        
        

