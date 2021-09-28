
import pytesseract
import cv2
import matplotlib.pyplot as plt
from PIL import Image
import re
import os
import time
def binarizing(img, threshold):
    """传入image对象进行灰度、二值处理"""
    img = img.convert("L")  # 转灰度

    pixdata = img.load()
    w, h = img.size
    # 遍历所有像素，大于阈值的为黑色
    for y in range(h):
        for x in range(w):
            if pixdata[x, y] < threshold:
                pixdata[x, y] = 0
            else:
                pixdata[x, y] = 255
    return img
def clear_border(img):
    pixdata = img.load()
    h, w = img.size
    for y in range(0, w):
        for x in range(0, h):
            if y < 2 or y > w - 2:
                pixdata[x, y] = 255
            if x < 2 or x > h -2:
                pixdata[x, y] = 255
    return img
def v_code(img):
    width, height = img.size
    print(img.size)
    # 1 二值化  调阈值
    b_img = binarizing(img, 140)
    b_img=clear_border(b_img)
    path = r'C:\Users\Administrator\Desktop\vcode\aaa.jpg'
    b_img.save(path)
    img2 = cv2.imread(path, 0)
    thresh=cv2.threshold(img2,0,255,cv2.THRESH_BINARY_INV+cv2.THRESH_OTSU)[1]
    data = pytesseract.image_to_string(thresh, lang="eng", config='--psm 6')
    data1=re.sub(r"\s*-*~*\?*¥*","",data)
    data2 = data1.replace("|","I")
    data3 = data2.replace("!", "I")
    print(data3)

if __name__ == '__main__':
  # Z=F
    for i in range(0,100):
        filename = str(i)+'.jpg'
        p = Image.open(filename)
        print(i)
        plt.figure("1")
        plt.imshow(p)
        plt.show()
        v_code(p)
        time.sleep(3)
