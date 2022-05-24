


import cv2
import numpy as np

img_path = r'C:\Users\evan\Documents\GitHub\OCR-Project\BAKR\Crop_IMG__1651113545.png'

img = cv2.imread(img_path)

gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
thresh_img = cv2.threshold(gray_img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]

cnts = cv2.findContours(thresh_img, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

boxes = list()

for i, cnt in enumerate(cnts):
    x,y,w,h = cv2.boundingRect(cnt)
    aspect_ratio = float(w)/h
    area = cv2.contourArea(cnt)
    rect_area = w*h
    extent = float(area)/rect_area
    if abs(aspect_ratio - 1) < 0.1 and extent > 0.7:
        print((x,y,w,h))
        boxes.append((x,y,w,h))