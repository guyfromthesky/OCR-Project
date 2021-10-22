from PIL import Image
import pytesseract
import cv2
import numpy as np
import os
image = ['Kr_1.jpg', 'Kr_2.jpg', 'Kr_3.jpg', 'Kr_4.jpg', 'Kr_5.jpg', 'Kr_6.jpg']

# If you don't have tesseract executable in your PATH, include the following:
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
_tess_data_path = r'"C:\Program Files\Tesseract-OCR\tessdata"'
tessdata_dir_config = '--psm 7 --tessdata-dir ' + _tess_data_path
print(tessdata_dir_config)

# Example tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract'

# List of available languages
print(pytesseract.get_languages(config=tessdata_dir_config))

def preprocess(img):
	img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
	#img = cv2.resize(img, None, fx=0.5, fy=0.5, interpolation=cv2.INTER_AREA)
	img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
	img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_LINEAR)
	#img = cv2.blur(img,(5,5))
	#img = cv2.GaussianBlur(img, (5, 5), 0)
	#img = cv2.medianBlur(img, 3)
	#img = cv2.bilateralFilter(img,9,75,75)
	return img

def image_smoothening(img):
	BINARY_THREHOLD = 100
	ret1, th1 = cv2.threshold(img, BINARY_THREHOLD, 255, cv2.THRESH_BINARY)
	ret2, th2 = cv2.threshold(th1, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
	#blur = cv2.GaussianBlur(th2, (1, 1), 0)
	ret3, th3 = cv2.threshold(th2, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
	return th3

def remove_noise_and_smooth(img):
	filtered = cv2.adaptiveThreshold(img.astype(np.uint8), 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 41, 3)
	kernel = np.ones((1, 1), np.uint8)
	opening = cv2.morphologyEx(filtered, cv2.MORPH_OPEN, kernel)
	closing = cv2.morphologyEx(opening, cv2.MORPH_CLOSE, kernel)
	img = image_smoothening(closing)
	
	#or_image = cv2.bitwise_or(img, closing)
	return img	

for img in image:
	sourcename = os.path.splitext(img)[0]
	img = cv2.imread(img)
	img = preprocess(img)
	img = image_smoothening(img)
	#cv2.imwrite(sourcename + '_processed.jpg', img)
	ocr = pytesseract.image_to_string(img, lang = 'kor', config=tessdata_dir_config)
	ocr = ocr.replace("\n","")
	ocr = ocr.replace("\r","")
	ocr = ocr.replace("\x0c","")
	print(ocr)
