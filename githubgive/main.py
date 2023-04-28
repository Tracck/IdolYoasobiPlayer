import cv2
import time
import xlwings as xw
import numpy as np
import pyautogui
import pygetwindow as gw
from multiprocessing import Process
from xlwings.utils import rgb_to_int

A_UPPERCASE = ord('A')
ALPHABET_SIZE = 26

def _decompose(number):
    """Generate digits from `number` in base alphabet, least significants
    bits first.

    Since A is 1 rather than 0 in base alphabet, we are dealing with
    `number - 1` at each iteration to be able to extract the proper digits.
    """

    while number:
        number, remainder = divmod(number - 1, ALPHABET_SIZE)
        yield remainder


def b10to26(number):
    """Convert a decimal number to its base alphabet representation"""

    return ''.join(
            chr(A_UPPERCASE + part)
            for part in _decompose(number)
    )[::-1]
    
def rgb_to_hex(r, g, b):
    return "#{:02x}{:02x}{:02x}".format(r, g, b)

def hex_to_rgb(value):
    value = value.lstrip('#')
    lv = len(value)
    return tuple(int(value[i:i + lv // 3], 16) for i in range(0, lv, lv // 3))

video = cv2.VideoCapture("./IdolYoasobi144p.mp4")
while not video.isOpened():
    video = cv2.VideoCapture("./IdolYoasobi144p.mp4")
    cv2.waitKey(1000)

# .xlsm to save macros
wb = xw.Book("YourExcelFileName.xlsm")
sheet = wb.sheets["YourExcelSheet"]
clearFormats = wb.macro("ClearFormats")

applyColors = np.empty(15, dtype=object)
applyColors[0] = wb.macro("ApplyColor1")
applyColors[1] = wb.macro("ApplyColor2")
applyColors[2] = wb.macro("ApplyColor3")
applyColors[3] = wb.macro("ApplyColor4")
applyColors[4] = wb.macro("ApplyColor5")
applyColors[5] = wb.macro("ApplyColor6")
applyColors[6] = wb.macro("ApplyColor7")
applyColors[7] = wb.macro("ApplyColor8")
applyColors[8] = wb.macro("ApplyColor9")
applyColors[9] = wb.macro("ApplyColor10")
applyColors[10] = wb.macro("ApplyColor11")
applyColors[11] = wb.macro("ApplyColor12")
applyColors[12] = wb.macro("ApplyColor13")
applyColors[13] = wb.macro("ApplyColor14")
applyColors[14] = wb.macro("ApplyColor15")

def task(start, frame):
    arrColorRange = {}
    
    for i in range(start - 24, start):
        line = frame[i]
        
        color1 = line[0]
        name1 = b10to26(1) + str(i + 1)
        name2 = name1
        for j in range(256):
            color2 = line[j]
            
            if (not set(color1).issubset(set(color2))) or (j == 255):
                hexColor = rgb_to_hex(color1[0], color1[1], color1[2])
                # 21 because I don't want an index error when it try to access arr[21] to check if there's any element
                if (arrColorRange.get(hexColor) is None):
                    arrColorRange[hexColor] = [[], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], []]
                    
                for i2 in range(20):
                    if (len(arrColorRange[hexColor][i2]) < 20):
                        arrColorRange[hexColor][i2].append(name1 + ":" + name2)
                        break
                    
                name1 = b10to26(j + 1) + str(i + 1)
                color1 = color2
            
            name2 = b10to26(j + 1) + str(i + 1)
            
    for colorHex, nameRange in arrColorRange.items():
        # I know it's r, g, b but apparently the conversion function does it like this
        b, g, r = hex_to_rgb(colorHex)
        names = []
        m = 0
        
        for i in range(16):
            if (len(nameRange[i]) == 0):
                m = i
                break 
            
            names.append(",".join(nameRange[i]))
        chosenApplyColor = applyColors[m - 1]
        
        # Hey guys at least I didn't use if statements
        match m:
            case 1:
                chosenApplyColor(names[0], r, g, b)
            case 2:
                n1, n2 = tuple(names)
                chosenApplyColor(n1, n2, r, g, b)
            case 3:
                n1, n2, n3 = tuple(names)
                chosenApplyColor(n1, n2, n3, r, g, b)
            case 4:
                n1, n2, n3, n4 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, r, g, b)
            case 5:
                n1, n2, n3, n4, n5 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, n5, r, g, b)
            case 6:
                n1, n2, n3, n4, n5, n6 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, n5, n6, r, g, b)
            case 7:
                n1, n2, n3, n4, n5, n6, n7 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, n5, n6, n7, r, g, b)
            case 8:
                n1, n2, n3, n4, n5, n6, n7, n8 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, n5, n6, n7, n8, r, g, b)
            case 9:
                n1, n2, n3, n4, n5, n6, n7, n8, n9 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, n5, n6, n7, n8, n9, r, g, b)
            case 10:
                n1, n2, n3, n4, n5, n6, n7, n8, n9, n10 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, r, g, b)
            case 11:
                n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, r, g, b)
            case 12:
                n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12, r, g, b)
            case 13:
                n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12, n13 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12, n13, r, g, b)
            case 14:
                n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12, n13, n14 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12, n13, n14, r, g, b)
            case 15:
                n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12, n13, n14, n15 = tuple(names)
                chosenApplyColor(n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12, n13, n14, n15, r, g, b)
    
# The frame of the video you want to start with
pos_frame = 1
video.set(cv2.CAP_PROP_POS_FRAMES, pos_frame)
while True:
    flag, frame = video.read()
    if flag:
        pos_frame += 1
    else:
        # The next frame is not ready, so we try to read it again
        video.set(cv2.CAP_PROP_POS_FRAMES, pos_frame - 1)
        # It is better to wait for a while for the next frame to be ready
        cv2.waitKey(1000)

    # The frame you want to stop + 2 (eg video has 200 frames, this number has to be 202)
    if pos_frame == 202:
        # If the number of captured frames is equal to the total number of frames,
        # we stop
        break
                
    # Some multi threading, so it renders faster
    if __name__ == '__main__':
        processes = [Process(target=task, args=(i * 24, frame,)) for i in range(1, 7)]

        for process in processes:
            process.start()
            
        for process in processes:
            process.join()
            
        time.sleep(2)
        pyautogui.screenshot().save("PathToSaveFrame:\FrameName.png")
        time.sleep(1)
        clearFormats()
        
        if (pos_frame % 3 == 0):
            wb.app.quit()
            time.sleep(1)
            wb = xw.Book("YourExcelFile.xlsm")
            sheet = wb.sheets["YourExcelSheet"]
            # Reinitialize
            clearFormats = wb.macro("ClearFormats")

            applyColors[0] = wb.macro("ApplyColor1")
            applyColors[1] = wb.macro("ApplyColor2")
            applyColors[2] = wb.macro("ApplyColor3")
            applyColors[3] = wb.macro("ApplyColor4")
            applyColors[4] = wb.macro("ApplyColor5")
            applyColors[5] = wb.macro("ApplyColor6")
            applyColors[6] = wb.macro("ApplyColor7")
            applyColors[7] = wb.macro("ApplyColor8")
            applyColors[8] = wb.macro("ApplyColor9")
            applyColors[9] = wb.macro("ApplyColor10")
            applyColors[10] = wb.macro("ApplyColor11")
            applyColors[11] = wb.macro("ApplyColor12")
            applyColors[12] = wb.macro("ApplyColor13")
            applyColors[13] = wb.macro("ApplyColor14")
            applyColors[14] = wb.macro("ApplyColor15")
            time.sleep(1)
    
            win = gw.getWindowsWithTitle("YourExcelFileName.xlsm - Excel")[0]
            win.maximize()