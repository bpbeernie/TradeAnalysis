import os
import glob
from datetime import date
from default import Constants as const

today = date.today()
todayString = today.strftime("%Y%m%d")

for xlsxfile in glob.glob(const.PATH + f'/*.xlsx'):
    originalFile = xlsxfile.replace("\\", "/")
    archivePath = f'{const.PATH}/Archive/{todayString}'
    if not os.path.exists(archivePath):
        os.makedirs(archivePath)
    
    os.rename(originalFile, f'{archivePath}/{os.path.basename(xlsxfile)}')
