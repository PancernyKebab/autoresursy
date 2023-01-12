import os
p=os.path.join(os.path.expanduser('~'),'Documents',"Autoresursy")

try:
    os.mkdir(p)
except FileExistsError:
    print("folder juz jest")
    pass
