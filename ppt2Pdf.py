import glob
import os
import tqdm
import win32com.client
from pathlib import Path

ppttoPDF = 32  # I forget what is this
PDFLOCATION = "./PDF"


def getFolder():
    srcStr = r"""
Please input the full destination folder(ex. C:\Users\abcd1\Class\Embedded System):
    >"""
    srcDir = input(srcStr)
    print()
    return srcDir  # r'C:\Users\abcd1\Documents\My Files\大二下\星期四 計算機組織\Handouts'


def save2PDF(file, endwith):
    savingfile = pdfPath + "\\" + file.split("\\")[-1]

    try:
        deck = powerpoint.Presentations.Open(file)
        deck.SaveAs(savingfile[: -len(endwith)], ppttoPDF)
        deck.Close()
        # delete file
        # os.remove(file)

    except BaseException as e:  # catch pywintypes.error(inherit from BaseException)
        print(e.args)
        pass


def pptChanger(files1, files2):
    for f in tqdm.tqdm(files1):  # ppt
        print()
        print(f)

        save2PDF(f, ".ppt")

    for f in tqdm.tqdm(files2):  # pptx
        print()
        print(f)

        save2PDF(f, ".pptx")


def createPDFfolder(dir):
    Path(dir).mkdir(parents=True, exist_ok=True)


if __name__ == "__main__":
    srcPath = getFolder()
    pdfPath = srcPath + "\\" + PDFLOCATION[2:]
    createPDFfolder(pdfPath)

    files1 = [f for f in glob.glob(srcPath + "**/*.ppt", recursive=True)]  # ppt
    files2 = [f for f in glob.glob(srcPath + "**/*.pptx", recursive=True)]  # pptx

    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    pptChanger(files1, files2)
    powerpoint.Quit()
