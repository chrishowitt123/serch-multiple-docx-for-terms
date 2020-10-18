import glob
from pathlib import Path
import os, sys
from docx import *
import docx2txt

folder = r'C:\Users\chris\Documents\Transgola\Clients\PROJECTS\2020\375091020_TM_JTI\Translation\MY COPY\Perfectit\Final\Splits'

orignalExt = '\*.docx'

orignalFiles = folder + orignalExt

sourceFiles = []

for filepath in glob.iglob(orignalFiles):
    sourceFiles.append(filepath)
    
newFileNames = []

for f in sourceFiles:
    base = os.path.basename(f)
    os.path.splitext(base)
    newFileName = os.path.splitext(base)[0]
    newFileNames.append(newFileName)
    
    
text = []

for file in sourceFiles:
    
    text.append(docx2txt.process(file))
    
search = ['KAs' , 'CA']

fnt = dict(zip(newFileNames, text))


for s in search:
    for k, v in fnt.items():
        if s in v:
            print(k + '   ------>   ' +  s)
