import cv2
import numpy as np



img_path = r'D:\App\OCR Project\Fellow Gacha\Screenshot_20211025-145143_V4.jpg'
img =  cv2.imread(img_path)
(_h, _w) = img.shape[:2]
img = cv2.resize(img,(int(_w*0.5),int(_h*0.5)))
edges = cv2.Canny(img,100,200)
kernal = np.ones((2,2),np.uint8)
dilation = cv2.dilate(edges, kernal , iterations=2)
bilateral = cv2.bilateralFilter(dilation,9,75,75)
contours, hireracy = cv2.findContours(bilateral,cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)
for i,contour in enumerate(contours):
    approx = cv2.approxPolyDP(contour, 0.01*cv2.arcLength(contour,True),True)   
    if len(approx) ==4:
        X,Y,W,H = cv2.boundingRect(approx)
        aspectratio = float(W)/H
        if aspectratio >=1.2 :
            box = cv2.rectangle(img, (X,Y), (X+W,Y+H), (0,0,255), 2)
            cropped = img[Y: Y+H, X: X+W]
            cv2.drawContours(img, [approx], 0, (0,255,0),5)
            x = approx.ravel()[0]
            y = approx.ravel()[1]
            cv2.putText(img, "rectangle"+str(i), (x,y),cv2.FONT_HERSHEY_COMPLEX, 0.5, (0,255,0))
cv2.imshow("image",img)
cv2.waitKey(0)
cv2.destroyAllWindows()