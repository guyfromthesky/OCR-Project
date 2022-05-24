import cv2
import numpy as np



img_path = r'C:\Users\evan\Documents\GitHub\OCR-Project\BAKR\Crop_IMG__1651113545.png'

# Let's load a simple image with 3 black squares
image = cv2.imread(img_path)
cv2.waitKey(0)

# Grayscale
gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

# Find Canny edges
edged = cv2.Canny(gray, 20, 20)
cv2.waitKey(0)

# Finding Contours
# Use a copy of the image e.g. edged.copy()
# since findContours alters the image
contours, hierarchy = cv2.findContours(edged,
	cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE)

h, w = edged.shape[:2]
mask = np.zeros((h+2, w+2), np.uint8)

cv2.floodFill(edged, mask, (0,0), 255);
im_floodfill_inv = cv2.bitwise_not(edged)


cv2.imshow('Canny Edges After Contouring', edged)
cv2.waitKey(0)

print("Number of Contours found = " + str(len(contours)))

# Draw all contours
# -1 signifies drawing all contours
boxes = []
for i,contour in enumerate(contours):
	approx = cv2.approxPolyDP(contour, 0.01*cv2.arcLength(contour,True),True)   
	if len(approx) ==4:
		X,Y,W,H = cv2.boundingRect(approx)
		box = {'x':X,'y':Y,'w':W,'h':H}
		boxes.append(box)
		#boxes.append(X,Y,W,H)
		#aspectratio = float(W)/H
		#if aspectratio <=1.2 :
		#	box = cv2.rectangle(image, (X,Y), (X+W,Y+H), (0,0,255), 2)
		#	cropped = image[Y: Y+H, X: X+W]
		#	#cv2.drawContours(image, [approx], 0, (0,255,0),5)
		#	x = approx.ravel()[0]
		#	y = approx.ravel()[1]
		#	#cv2.putText(image, "rectangle"+str(i), (x,y),cv2.FONT_HERSHEY_COMPLEX, 0.5, (0,255,0))

# Recalculate the box w and h:
total_boxes = len(boxes)
total_w = 0
total_h = 0
x_list = []
y_list = []
for box in boxes:
	total_w += box['w']
	total_h += box['h']
	x_list.append(box['x'])
	y_list.append(box['y'])
avg_w = int(total_w/total_boxes)
avg_h = int(total_h/total_boxes)

for box in boxes:
	box['w'] = avg_w
	box['h'] = avg_h

# Update all x value in boxes if it is has similar value with other boxes
for i,box in enumerate(boxes):
	for j,box2 in enumerate(boxes):
		if i != j:
			one_percent_x = int(box['x']*0.01)
			one_percent_y = int(box['y']*0.01)
			if abs(box['x'] - box2['x'] ) < one_percent_x:
				box['x'] = box2['x']
			if abs(box['y'] - box2['y'] ) < one_percent_y:
				box['y'] = box2['y']

for box in boxes:	
	print(box)
	cv2.rectangle(image, (box['x'],box['y']), (box['x']+box['w'],box['y']+box['h']), (0,255,0), 1)

cv2.imshow("image",image)
cv2.waitKey(0)
cv2.destroyAllWindows()

