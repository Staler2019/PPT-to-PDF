import glob

ppttoPDF = 32  # I forget what is this


def getFolder():
    srcStr = r"""
Please input the full destination folder(ex. C:\Users\abcd1\Class\Embedded System):
    >"""
    srcDir = input(srcStr)
    print()
    # r'C:\Users\abcd1\Documents\My Files\大二下\星期四 計算機組織\Handouts'
    return srcDir


def pptChanger(files1, files2):
    import os
    import tqdm
    import win32com.client

    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    for f in tqdm.tqdm(files1):  # ppt
        print()
        print(f)

        try:
            # kernel
            deck = powerpoint.Presentations.Open(f)
            deck.SaveAs(f[:-4], ppttoPDF)
            deck.Close()
            os.remove(f)

        except BaseException as e:  # catch pywintypes.error(inherit from BaseException)
            print(e.args)
            pass

    for f in tqdm.tqdm(files2):  # pptx
        print()
        print(f)

        try:
            # kernel
            deck = powerpoint.Presentations.Open(f)
            deck.SaveAs(f[:-5], ppttoPDF)
            deck.Close()
            os.remove(f)

        except BaseException as e:  # to catch pywintypes.error
            print(e.args)
            pass

    powerpoint.Quit()


if __name__ == "__main__":
    srcPath = getFolder()

    files1 = [f for f in glob.glob(srcPath + "**/*.ppt", recursive=True)]  # ppt
    files2 = [f for f in glob.glob(srcPath + "**/*.pptx", recursive=True)]  # pptx

    pptChanger(files1, files2)
