import os
from pathlib import Path
import win32com.client

import glob
import tqdm


ppttoPDF = 32
PDF_LOC = "PDF"


def getFolder():
    srcStr = r"""
Please input the full destination folder(ex. C:\Users\abcd1\Class\Embedded System):
    >"""
    srcDir = input(srcStr)
    print()
    return srcDir  # r'C:\Users\abcd1\Documents\My Files\大二下\星期四 計算機組織\Handouts'


def save2pdf(f):
    fileName = os.path.splitext(f.split("\\")[-1])[0] + ".pdf"
    fileSaveDir = os.path.join(os.path.dirname(f), PDF_LOC)
    saveFileName =  os.path.join(fileSaveDir, fileName)

    Path(fileSaveDir).mkdir(parents=True, exist_ok=True)

    print("to", saveFileName)

    try:
        # Load presentation
        deck = powerpoint.Presentations.Open(f)
        # Convert PPTX to PDF
        deck.SaveAs(saveFileName, ppttoPDF)
        deck.Close()

        # delete file
        # os.remove(file)

    except BaseException as e:  # catch pywintypes.error(inherit from BaseException)
        print(e.args)


if __name__ == "__main__":
    srcPath = getFolder()

    types = ("/**/*.ppt", "/**/*.pptx")
    files = []
    for ftype in types:
        files.extend(glob.glob(srcPath + ftype, recursive=True))

    powerpoint = win32com.client.Dispatch("Powerpoint.Application")

    for f in tqdm.tqdm(files):
        print()
        print(f)
        save2pdf(f)

    powerpoint.Quit()
