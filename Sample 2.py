import os
import cv2
import numpy as np
import pytesseract
import csv
import shutil

def Compare_2_Image(source_image_path, target_image_path):
	source_image = cv2.imread(source_image_path)
	source_image = cv2.cvtColor(source_image, cv2.COLOR_BGR2GRAY)	

	target_image = cv2.imread(target_image_path)
	target_image = cv2.cvtColor(target_image, cv2.COLOR_BGR2GRAY)	
	
	result = cv2.matchTemplate(source_image, target_image, cv2.TM_CCOEFF_NORMED)
	(_, maxVal, _, maxLoc) = cv2.minMaxLoc(result)
	
	if maxVal > 0.8:
		#print('maxVal', maxVal)
		return True
	else:
		print('source_image_path', source_image_path)
		print('target_image_path', target_image_path)
		print('maxVal', maxVal)
		return False

def Compare_All():
	path = r'C:\Users\evan\Documents\GitHub\OCR-Project\Test Image'
	
	_temp_image_files = os.listdir(path)
	all_images = []
	for image in _temp_image_files:
		image_path = path + '\\' + image
		if os.path.isfile(image_path):
			all_images.append(path + '\\' + image)
	unique = []
	count = {}
	for source_image in all_images:
		baseName = os.path.basename(source_image)
		if len(unique) == 0:
			count[baseName] = 1
			unique.append(source_image)
			print_result(source_image)
		else:
			result = False
			for target_image in unique:
				result = Compare_2_Image(source_image, target_image)
				if result == True:
					base_target = os.path.basename(target_image)
					count[base_target] += 1
					break
			if result == False:
				print('Append to unique:',source_image )
				count[baseName] = 1
				unique.append(source_image)
				print_result(source_image)
	print(count)
	
def print_result(path):
	unique_path = r'C:\Users\evan\Documents\GitHub\OCR-Project\Test Image\unique'
	unique_image = cv2.imread(path)
	baseName = os.path.basename(path)
	new_name = unique_path+'\\' + baseName
	cv2.imwrite(new_name, unique_image)

Compare_All()
