import os
import cv2
import numpy as np
import pytesseract
import csv

def Load_Image_by_Ratio(image_path, resolution):
	_img = cv2.imread(image_path)
	(_h, _w) = _img.shape[:2]
	_ratio = resolution / _h
	if _ratio != 1:
		width = int(_img.shape[1] * _ratio)
		height = int(_img.shape[0] * _ratio)
		dim = (width, height)
		_img = cv2.resize(_img, dim, interpolation = cv2.INTER_AREA)
	return _img

def Function_Pre_Processing_Image(img):
	img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
	#img = cv2.resize(img, None, fx=0.5, fy=0.5, interpolation=cv2.INTER_AREA)
	img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
	img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_LINEAR)
	Img = cv2.blur(img,(5,5))
	img = cv2.GaussianBlur(img, (5, 5), 0)
	img = cv2.medianBlur(img, 3)
	img = cv2.bilateralFilter(img,9,75,75)
	#img = cv2.threshold(img,127,255,cv2.THRESH_BINARY)
	img = image_smoothening(img)
	return	img

def image_smoothening(img):
	BINARY_THREHOLD = 100
	ret1, th = cv2.threshold(img, BINARY_THREHOLD, 255, cv2.THRESH_BINARY)
	ret2, th = cv2.threshold(th, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
	#th = cv2.GaussianBlur(th (1, 1), 0)
	ret3, th = cv2.threshold(th, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
	return th

def Get_Text(img, tess_language, tessdata_dir_config):
	ocr = pytesseract.image_to_string(img, lang = tess_language, config=tessdata_dir_config)
	ocr = ocr.replace("\n", "")
	ocr = ocr.replace("\r", "")  
	ocr = ocr.replace("\x0c", "") 
	return ocr

def Get_Text_From_Single_Image(tess_path, tess_language, advanced_tessdata_dir_config, input_image, ratio, scan_areas, result_file,):

	pytesseract.pytesseract.tesseract_cmd = tess_path
	_img = Load_Image_by_Ratio(input_image, ratio)
	_result = []
	_output_dir = os.path.dirname(result_file)
	baseName = os.path.basename(input_image)
	sourcename = os.path.splitext(baseName)[0]
	_area_count = 0
	for area in scan_areas:
		_area_count +=1
		imCrop = _img[int(area[1]):int(area[1]+area[3]), int(area[0]):int(area[0]+area[2])]
		imCrop = Function_Pre_Processing_Image(imCrop)
		_name = _output_dir + '\\' + sourcename + '_' + str(_area_count) + '.jpg'
		#cv2.imwrite(_name, imCrop)
		#cv2.imshow(_name, imCrop)
		#cv2.waitKey(2000)
		ocr = Get_Text(imCrop, tess_language, advanced_tessdata_dir_config)
		_result.append(ocr)
	print(_result)
	baseName = os.path.basename(input_image)
	file_name = os.path.splitext(baseName)[0]
	while True:
		try:
			with open(result_file, 'a', newline='', encoding='utf-8-sig') as csvfile:
				writer = csv.writer(csvfile)
				writer.writerow([file_name] + _result)
				break
		except PermissionError:
			continue

tess_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
tess_language = 'kor'
advanced_tessdata_dir_config = r'--psm 7 --tessdata-dir "C:\Program Files\Tesseract-OCR\tessdata"'
input_image = r'D:\App\OCR Project\Fellow Gacha\Screenshot_20211025-145143_V4.jpg'
ratio = 720
scan_areas = [
	[310, 310, 150,35], [517, 310, 150,35],
	[724, 310, 150,35], [931, 310, 150,35],
	[1138, 310, 150,35], [205, 535, 150,35],
	[412, 535, 150,35], [619, 535, 150,35],
	[826, 535, 150,35], [1033, 535, 150,35]
]
result_file = r'D:\App\OCR Project\Fellow Gacha\Scan_Result_1635154972\result.csv'

Get_Text_From_Single_Image(tess_path, tess_language, advanced_tessdata_dir_config, input_image, ratio, scan_areas, result_file)