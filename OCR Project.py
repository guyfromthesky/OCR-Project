#System variable and io handling
import os
import csv
from posixpath import basename
import sys

import configparser
#Regular expression handlings
import multiprocessing
from multiprocessing import Process , Queue, Manager
import queue 
import subprocess
#Get timestamp
import time
from datetime import datetime
#function difination
import unicodedata
from urllib.parse import urlparse

#GUI
from tkinter.ttk import Entry, Label, Treeview, Scrollbar, OptionMenu
from tkinter.ttk import Checkbutton, Button, Notebook, Radiobutton
from tkinter.ttk import Progressbar, Style

from tkinter import Tk, Frame
from tkinter import Menu, filedialog, messagebox
from tkinter import Text, colorchooser
from tkinter import IntVar, StringVar
from tkinter import W, E, S, N, END, RIGHT, HORIZONTAL, NO, CENTER
from tkinter import WORD, NORMAL, BOTTOM, X, TOP, BOTH, Y
from tkinter import DISABLED

from tkinter import scrolledtext 
from tkinter import simpledialog

import webbrowser

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.styles import Color
from openpyxl.styles import Color, PatternFill, Font

from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter, column_index_from_string

from libs.configmanager import ConfigLoader
from libs.version import get_version
from libs.tkinter_extension import AutocompleteCombobox

import cv2
import numpy as np

#from document_toolkit_function.py import *

DELAY1 = 20

ToolDisplayName = "OCR Project"
tool_name = 'ocr'
rev = 1000
a,b,c,d = list(str(rev))
VerNum = a + '.' + b + '.' + c + chr(int(d)+97)

version = ToolDisplayName  + " " +  VerNum 

#**********************************************************************************
# UI handle ***********************************************************************
#**********************************************************************************

class OCR_Project(Frame):
	def __init__(self, Root, Queue = None, Manager = None,):
		
		Frame.__init__(self, Root) 
		#super().__init__()
		self.parent = Root 

		# Queue
		self.Process_Queue = Queue['Process_Queue']
		self.Result_Queue = Queue['Result_Queue']
		self.Status_Queue = Queue['Status_Queue']
		self.Debug_Queue = Queue['Debug_Queue']
		self.Manager = Manager['Default_Manager']

		self.Options = {}

		# XLSX Optmizer
		self.Optimize_Folder = ""
		self.Optimize_FileList = ""
		# XLSX Comparision
		self.Compare_Folder_Old = ""
		self.Compare_File_List_Old = ""
		self.Compare_Folder_New = ""
		self.Compare_File_List_New = ""

		# UI Variable
		self.Button_Width_Full = 20
		self.Button_Width_Half = 15
		
		self.PadX_Half = 5
		self.PadX_Full = 10
		self.PadY_Half = 5
		self.PadY_Full = 10
		self.StatusLength = 120
		self.AppLanguage = 'en'

		self.OCR_File_Path = None

		self.Path_Size = 80

		self.init_App_Setting()
		
		self.App_LanguagePack = {}
		

		if self.AppLanguage != 'kr':
			from libs.languagepack import LanguagePackEN as LanguagePack
		else:
			from libs.languagepack import LanguagePackKR as LanguagePack

		self.LanguagePack = LanguagePack

		# Init function

		self.parent.resizable(False, False)
		self.parent.title(version)
		# Creating Menubar 
		
		#**************New row#**************#
		self.Notice = StringVar()
		self.Debug = StringVar()
		self.Progress = StringVar()
	
		self.basePath = os.path.abspath(os.path.dirname(sys.argv[0]))
		self.ExceptionPath = self.basePath + "\\Exception.xlsx"
		try:
			self.ExceptionList = self.ImportException(self.ExceptionPath)
			print('My exception list: ', self.ExceptionList)
		except:
			self.ExceptionList = []
		
		#Generate UI
		self.Generate_Menu_UI()
		self.Generate_Tab_UI()
		self.init_UI()
		



	# UI init
	def init_UI(self):
	
		self.Generate_OCR_Tool_UI(self.OCR_TOOL)

	def Generate_Menu_UI(self):
		menubar = Menu(self.parent) 
		# Adding File Menu and commands 
		'''
		file = Menu(menubar, tearoff = 0)
		
		# Adding Load Menu 
		menubar.add_cascade(label =  self.LanguagePack.Menu['File'], menu = file) 
		file.add_command(label =  self.LanguagePack.Menu['LoadTM'], command = self.Menu_Function_Select_TM) 
		file.add_separator() 
		file.add_command(label =  self.LanguagePack.Menu['CreateTM'], command = self.Menu_Function_Create_TM)
		file.add_separator() 
		file.add_command(label =  self.LanguagePack.Menu['Exit'], command = self.parent.destroy) 
		'''
		# Adding Help Menu
		help_ = Menu(menubar, tearoff = 0) 
		menubar.add_cascade(label =  self.LanguagePack.Menu['Help'], menu = help_) 
		help_.add_command(label =  self.LanguagePack.Menu['GuideLine'], command = self.Menu_Function_Open_Main_Guideline) 
		help_.add_separator()
		help_.add_command(label =  self.LanguagePack.Menu['About'], command = self.Menu_Function_About) 
		self.parent.config(menu = menubar)

		# Adding Help Menu
		language = Menu(menubar, tearoff = 0) 
		menubar.add_cascade(label =  self.LanguagePack.Menu['Language'], menu = language) 
		language.add_command(label =  self.LanguagePack.Menu['Hangul'], command = self.SetLanguageKorean) 
		language.add_command(label =  self.LanguagePack.Menu['English'], command = self.SetLanguageEnglish) 
		self.parent.config(menu = menubar) 	

	def Generate_Tab_UI(self):
		
		self.TAB_CONTROL = Notebook(self.parent)

		self.OCR_TOOL = Frame(self.TAB_CONTROL)
		#self.TAB_CONTROL.add(self.AutoTest, text= self.LanguagePack.Tab['AutomationTest'])
		self.TAB_CONTROL.add(self.OCR_TOOL, text= 'OCR Project')
		
		self.TAB_CONTROL.pack(expand=1, fill="both")
		return

	#STABLE
	def Generate_OCR_Tool_UI(self, Tab):
		
		Row=1
		self.Str_OCR_Image_Path = StringVar()
		Label(Tab, text=  self.LanguagePack.Label['ImageSource']).grid(row=Row, column=1, padx=5, pady=5, sticky= W)
		self.Entry_Old_File_Path = Entry(Tab,width = self.Path_Size, state="readonly", textvariable=self.Str_OCR_Image_Path)
		self.Entry_Old_File_Path.grid(row=Row, column=2, columnspan=7, padx=5, pady=5, sticky=E+W)
		 
		Btn_Browse_Image = Button(Tab, width = self.Button_Width_Half, text=  self.LanguagePack.Button['Browse'], command= self.Btn_OCR_Browse_Image_Data)
		Btn_Browse_Image.grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=W)

		Row+=1
		self.Str_OCR_Config_Path = StringVar()
		Label(Tab, text= self.LanguagePack.Label['ScanConfig']).grid(row=Row, column=1, padx=5, pady=5, sticky= W)
		self.Entry_Old_File_Path = Entry(Tab,width = self.Path_Size, state="readonly", textvariable=self.Str_OCR_Config_Path)
		self.Entry_Old_File_Path.grid(row=Row, column=2, columnspan=7, padx=5, pady=5, sticky=E+W)
		
		Btn_Browse_Setting = Button(Tab, width = self.Button_Width_Half, text=  self.LanguagePack.Button['Browse'], command= self.Btn_OCR_Browse_Config_File)
		Btn_Browse_Setting.grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=E)
		
		Row+=1
		Label(Tab, text= self.LanguagePack.Label['CenterX']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.Str_CenterX = Text(Tab, width = 10, height=1) #
		self.Str_CenterX.grid(row=Row, column=2, padx=5, pady=5, sticky=W)

		self.Str_CenterX.bind("<Tab>", self.entry_next)	

		Label(Tab, text= self.LanguagePack.Label['CenterY']).grid(row=Row, column=3, pady=5, sticky=W)
		self.Str_CenterY = Text(Tab, width=10, height=1) #
		self.Str_CenterY.grid(row=Row, column=4, pady=5, sticky=W)
		self.Str_CenterY.bind("<Tab>", self.entry_next)	

		Btn_Input_Area = Button(Tab, width = self.Button_Width_Half, text= self.LanguagePack.Button['AddAreaWithText'], command= self.Btn_OCR_Input_Area)
		Btn_Input_Area.grid(row=Row, column=5, padx=5, pady=5, sticky=W)

		Label(Tab, text= self.LanguagePack.Label['BrowseType']).grid(row=Row, column=6, rowspan=2, pady=5, sticky=W)

		self.ocr_data_select_type = IntVar()
		Radiobutton(Tab, width= 15, text=  self.LanguagePack.Option['Folder'], value=1, variable=self.ocr_data_select_type).grid(row=Row, column=7,columnspan=2,padx=0, pady=5, sticky=W)
		self.ocr_data_select_type.set(1)


		

		#Btn_Input_Area = Button(Tab, width = self.Button_Width_Half, text= self.LanguagePack.Button['AddAreaWithText'], command= self.Btn_OCR_Input_Area)
		#Btn_Input_Area.grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=W)

		Row+=1
		Label(Tab, text= self.LanguagePack.Label['Height']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.Str_Height = Text(Tab, width=10, height=1) #
		self.Str_Height.grid(row=Row, column=2, padx=5, pady=5, sticky=W)
		self.Str_Height.bind("<Tab>", self.entry_next)	
	
		Label(Tab, text= self.LanguagePack.Label['Weight']).grid(row=Row, column=3, pady=5, sticky=W)
		self.Str_Weight = Text(Tab, width = 10, height=1) #
		self.Str_Weight.grid(row=Row, column=4, pady=5, sticky=W)
		self.Str_Weight.bind("<Tab>", self.entry_next)	

		self.Btn_Update_Area = Button(Tab, width = self.Button_Width_Half, text= self.LanguagePack.Button['SaveConfig'], command= self.Btn_OCR_Update_Area)
		self.Btn_Update_Area.grid(row=Row, column=5, padx=5, pady=5, sticky=W)
		self.Btn_Update_Area.configure(state=DISABLED)

		Radiobutton(Tab, width= 15, text= self.LanguagePack.Option['File'], value=2, variable=self.ocr_data_select_type).grid(row=Row, column=7,columnspan=2, padx=0, pady=5, sticky=W)
		
		
		Row+=1
		self.Treeview = Treeview(Tab)
		self.Focused_Item = None
		self.Treeview.grid(row=Row, column=1, columnspan=8, rowspan=5, padx=5, pady=5, sticky = N+S+W+E)
		verscrlbar = Scrollbar(Tab, orient ="vertical", command = self.Treeview.yview)
		self.Treeview.configure( yscrollcommand=verscrlbar.set)
	
		self.Treeview.Scrollable = True
		self.Treeview['columns'] = ('X', 'Y', 'W', 'H')

		self.Treeview.column('#0', width=0, stretch=NO)
		self.Treeview.column('X', anchor=CENTER, width=100)
		self.Treeview.column('Y', anchor=CENTER, width=100)
		self.Treeview.column('W', anchor=CENTER, width=100)
		self.Treeview.column('H', anchor=CENTER, width=100)

		self.Treeview.heading('#0', text='', anchor=CENTER)
		self.Treeview.heading('X', text='X', anchor=CENTER)
		self.Treeview.heading('Y', text='Y', anchor=CENTER)
		self.Treeview.heading('W', text='W', anchor=CENTER)
		self.Treeview.heading('H', text='H', anchor=CENTER)

		verscrlbar.grid(row=Row, column=8, rowspan=5,  sticky = N+S+E)
		Tab.grid_columnconfigure(11, weight=0, pad=0)
		styles = Style()
		styles.configure('Treeview',rowheight=15)

		self.Treeview.bind("<Delete>", self.delete_treeview_line)	
		self.Treeview.bind("<Double-1>", self.Treeview_OCR_Select_Row)	

		Btn_Select_Area = Button(Tab, width = self.Button_Width_Half, text= self.LanguagePack.Button['SelectArea'], command= self.Btn_OCR_Select_Area)
		Btn_Select_Area.grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=W)
		
		Row+=1
		
		Btn_Preview_Area = Button(Tab, width = self.Button_Width_Half, text= self.LanguagePack.Button['PreviewArea'], command= self.Btn_OCR_Preview_Areas)
		Btn_Preview_Area.grid(row=Row, column=9, columnspan=2,padx=5, pady=5, sticky=W)
		
		Row+=1

		Btn_Save_Setting = Button(Tab, width = self.Button_Width_Half, text=  self.LanguagePack.Button['SaveConfig'], command= self.Btn_OCR_Save_Config_File)
		Btn_Save_Setting.grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=W)

		Row+=1

		Btn_Open_Result = Button(Tab, width = self.Button_Width_Half, text= self.LanguagePack.Button['OpenOutput'], command= self.Open_OCR_Result_Folder)
		Btn_Open_Result.grid(row=Row, column=9, columnspan=2,padx=5, pady=5, sticky=W)

		Row+=1

		Btn_Execute = Button(Tab, width = self.Button_Width_Half, text= self.LanguagePack.Button['Scan'], command= self.Btn_OCR_Execute)
		Btn_Execute.grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=W)

		Row+=1
		Label(Tab, text= self.LanguagePack.Label['Debug']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.Debugger = scrolledtext.ScrolledText(Tab, width=110, height=5, undo=False, wrap=WORD, )
		self.Debugger.grid(row=Row, column=2, columnspan=8, padx=5, pady=5, sticky=W+E+N+S)

		Row+=1
		Label(Tab, text= self.LanguagePack.Label['Progress']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.progressbar = Progressbar(Tab, orient=HORIZONTAL, length=900,  mode='determinate')
		self.progressbar["maximum"] = 1000
		self.progressbar.grid(row=Row, column=2, columnspan=8, padx=5, pady=5, sticky=W)

		
		


###########################################################################################
# Treeview FUNCTION
###########################################################################################

	def delete_treeview_line(self, event):
		'''
		Function activate when select an entry from a Treeview and press Delete btn
		'''
		selected = self.Treeview.selection()
		to_remove = []
		for child_obj in selected:
			child = self.Treeview.item(child_obj)
			tm_index = child['values'][0]
			to_remove.append(tm_index)
			self.Treeview.delete(child_obj)

	# Obsoleted.
	def double_right_click_treeview(self, event):
		'''
		Function activate when double click an entry from Treeview
		'''
		focused = self.Treeview.focus()
		child = self.Treeview.item(focused)
		self.Debugger.insert("end", "\n")
		self.Debugger.insert("end", 'Korean: ' + str(child["text"]))
		self.Debugger.insert("end", "\n")
		self.Debugger.insert("end", 'English: ' + str(child["values"][0]))
		self.Debugger.yview(END)


	# Nam will check
	def load_tm_list(self):
		"""
		When clicking the [Load] button in TM Manager tab
		Display the pair languages in the text box.
		"""
		self.remove_treeview()
		
		_area_list = []

		for location in _area_list:	
			try:
				self.Treeview.insert('', 'end', text= '', values=( str(location['index']), str(location['x']), str(location['y']), str(location['h']), str(location['w'])))
			except:
				pass

	def add_treeview_row(self, location):
		'''
		Add a row to the current Treeview
		'''
		self.Treeview.insert('', 'end', text= '', values=(str(location[0]), str(location[1]), str(location[2]), str(location[3])))

	def remove_treeview(self):
		for i in self.Treeview.get_children():
			self.Treeview.delete(i)

###########################################################################################
# MENU FUNCTION
###########################################################################################

	def Menu_Function_About(self):
		messagebox.showinfo("About....", "Creator: Evan")

	def Show_Error_Message(self, ErrorText):
		messagebox.showinfo('Error...', ErrorText)	

	def SaveAppLanguage(self, language):
		self.Write_Debug(self.LanguagePack.ToolTips['AppLanuageUpdate'] + " "+ language) 
		self.AppConfig.Save_Config(self.AppConfig.Ocr_Tool_Config_Path, 'OCR_TOOL', 'app_lang', language)

	def SetLanguageKorean(self):
		self.AppLanguage = 'kr'
		self.SaveAppLanguage(self.AppLanguage)
	
	def SetLanguageEnglish(self):
		self.AppLanguage = 'en'
		self.SaveAppLanguage(self.AppLanguage)

	def Function_Correct_Path(self, path):
		return str(path).replace('/', '\\')
	
	def Menu_Function_Open_Main_Guideline(self):
		webbrowser.open_new(r"https://confluence.nexon.com/display/NWMQA/OCR+Tool")

	def onExit(self):
		self.quit()

	def init_App_Setting(self):

		self.DB_Path = StringVar()
	
		self.CurrentDataSource = StringVar()
		self.Notice = StringVar()

		self.AppConfig = ConfigLoader()
		self.Configuration = self.AppConfig.Config
		self.AppLanguage  = self.Configuration['OCR_TOOL']['app_lang']
		
		db_path = self.Configuration['OCR_TOOL']['db_file']
		self.DB_Path.set(db_path)

	def SaveSetting(self):

		print('Save setting')
		return

	def Btn_Select_DB_Path(self):
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectDB'],filetypes = (("JSON files","*.xlsx" ), ), )	
		if filename != "":
			db_path = self.CorrectPath(filename)
			self.AppConfig.Save_Config(self.AppConfig.Ocr_Tool_Config_Path, 'OCR_TOOL', 'db_file', db_path, True)
		else:
			self.Write_Debug("No file is selected")

###########################################################################################
# General functions
###########################################################################################

	def CorrectPath(self, path):
		if sys.platform.startswith('win'):
			return str(path).replace('/', '\\')
		else:
			return str(path).replace('\\', '//')
	
	def CorrectExt(self, path, ext):
		if path != None and ext != None:
			Outputdir = os.path.dirname(path)
			baseName = os.path.basename(path)
			sourcename = os.path.splitext(baseName)[0]
			newPath = self.CorrectPath(Outputdir + '/'+ sourcename + '.' + ext)
			return newPath

	def Write_Debug(self, text):
		'''
		Function write the text to debugger box and move to the end of the box
		'''
		self.Debugger.insert("end", "\n")
		self.Debugger.insert("end", str(text))

		self.Debugger.yview(END)		

	def entry_next(self, event):
		event.widget.tk_focusNext().focus()
		return("break")

###########################################################################################
# OCR
###########################################################################################
	
	def Btn_OCR_Select_Background_Colour(self):
		colorStr, self.Background_Color = colorchooser.askcolor(parent=self, title='Select Colour')
		
		if self.Background_Color == None:
			self.Show_Error_Message('Set colour as defalt colour (Yellow)')
			self.Background_Color = 'ffff00'
		else:
			self.Background_Color = self.Background_Color.replace('#', '')
		#print(colorStr)
		#print(self.BackgroundColor)
		return
	
	def Btn_OCR_Select_Font_Colour(self):
		colorStr, self.Font_Color = colorchooser.askcolor(parent=self, title='Select Colour')
		
		
		if self.Font_Color == None:
			self.Show_Error_Message('Set colour as defalt colour (Yellow)')
			self.Font_Color = 'FF0000'
		else:
			self.Font_Color = self.Font_Color.replace('#', '')
		#print(colorStr)
		#print(self.BackgroundColor)
		return

	def Btn_OCR_Browse_Config_File(self):
		
		config_file = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'], filetypes = (("Config files", "*.csv *.xlsx"), ), multiple = False)	
		
		if os.path.isfile(config_file) != "":
			print('config_file', config_file)
			self.Str_OCR_Config_Path.set(config_file)
			self.remove_treeview()
			with open(config_file, newline='', encoding='utf-8') as csvfile:
				reader = csv.DictReader(csvfile)
				for location in reader:	
					self.Treeview.insert('', 'end', text= '', values=(str(location['x']), str(location['y']), str(location['w']), str(location['h'])))
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])

	def Btn_OCR_Save_Config_File(self):
		'''
		Save all added scan areas to csv file.
		'''
		filename = filedialog.asksaveasfilename(title = "Select file", filetypes = (("Scan Config", "*.csv"),),)
		filename = self.CorrectExt(filename, "csv")
		if filename == "":
			return
		else:
			with open(filename, 'w', newline='') as csvfile:
				fieldnames = ['x', 'y', 'w', 'h']
				writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
				writer.writeheader()
				for row in self.Treeview.get_children():
					child = self.Treeview.item(row)
					writer.writerow({'x': child["values"][0], 'y': child["values"][1], 'w': child["values"][2], 'h': child["values"][3]})
			
	def Btn_OCR_Browse_Image_Data(self):
		_select_type = self.ocr_data_select_type.get()
		if _select_type == 1:
			self.Btn_OCR_Browse_Image_Folder()
		else:
			self.Btn_OCR_Browse_Image_Files()

	def Btn_OCR_Browse_Image_Folder(self):
			
		folder_name = filedialog.askdirectory(title =  self.LanguagePack.ToolTips['SelectSource'],)	
		if folder_name != "":
			_temp_text_files = os.listdir(folder_name)
			self.OCR_File_Path = []
			for file in _temp_text_files:
				file_path = folder_name + '/' + file
				if os.path.isfile(file_path):
					baseName = os.path.basename(file_path)
					sourcename, ext = os.path.splitext(baseName)
					if 'jpg' in ext or 'jpeg' in ext:
						self.OCR_File_Path.append(file_path)

			self.Str_OCR_Image_Path.set(str(self.OCR_File_Path[0]))

			self.Write_Debug(self.LanguagePack.ToolTips['DataSelected'] + ": " + str(len(self.OCR_File_Path)))
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])

	def Btn_OCR_Browse_Image_Files(self):
			
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'], filetypes = (("Image files", "*.jpg *.jpeg"), ), multiple = True)	
		if filename != "":
			self.OCR_File_Path = list(filename)
			self.Str_OCR_Image_Path.set(str(self.OCR_File_Path[0]))
			
			self.Write_Debug(self.LanguagePack.ToolTips['DataSelected'] + ": " + str(len(self.OCR_File_Path)))
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])

	def Btn_OCR_Browse_DB_File(self):
			
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'], filetypes = (("Workbook files", "*.xlsx *.xlsm"), ), multiple = False)	
		if filename != "":
			self.BadWord_DB_Path = filename
			self.Str_BadWord_DB_Path.set(filename)
			self.Write_Debug(self.LanguagePack.ToolTips['DetaSelected'])	
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return
		
	def Btn_OCR_Select_Area(self):
		if self.OCR_File_Path != None:
			if os.path.isfile(self.OCR_File_Path[0]):
				im = cv2.imread(self.OCR_File_Path[0])
				(_h, _w) = im.shape[:2]
				ratio = 720 / _h
				if ratio != 1:
					width = int(im.shape[1] * ratio)
					height = int(im.shape[0] * ratio)
					dim = (width, height)
					im = cv2.resize(im, dim, interpolation = cv2.INTER_AREA)
				# Select ROI
				# Select multiple rectanglesx
				for row in self.Treeview.get_children():
					child = self.Treeview.item(row)
					im = cv2.rectangle(im, (child["values"][0], child["values"][1]), (child["values"][0] + child["values"][2], child["values"][1] + child["values"][3]), (255,0,0), 2)

				location = cv2.selectROI("Sekect scan area", im, showCrosshair=False,fromCenter=False)
				cv2.destroyAllWindows() 
				self.Treeview.insert('', 'end', text= '', values=(str(location[0]), str(location[1]), str(location[2]), str(location[3])))
			else:
				self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])		
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])	
		
				

	def Btn_OCR_Input_Area(self):
		_x = self.Str_CenterX.get("1.0", END).replace('\n', '')
		_y = self.Str_CenterY.get("1.0", END).replace('\n', '')
		_w = self.Str_Weight.get("1.0", END).replace('\n', '')
		_h = self.Str_Height.get("1.0", END).replace('\n', '')
		self.Treeview.insert('', 'end', text= '', values=(str(_x), str(_y), str(_w), str(_h)))
	
	def Treeview_OCR_Select_Row(self, event):
		'''
		Function activate when double click an entry from Treeview
		'''
		self.Focused_Item = self.Treeview.focus()
		child = self.Treeview.item(self.Focused_Item)
		self.Btn_Update_Area.configure(state=NORMAL)
		self.Str_CenterX.delete("1.0", END)
		self.Str_CenterX.insert("end", child["values"][0])

		self.Str_CenterY.delete("1.0", END)
		self.Str_CenterY.insert("end", child["values"][1])

		self.Str_Weight.delete("1.0", END)
		self.Str_Weight.insert("end", child["values"][2])

		self.Str_Height.delete("1.0", END)
		self.Str_Height.insert("end", child["values"][3])

		return

	def Btn_OCR_Update_Area(self):

		if self.Focused_Item != None:
			_x = self.Str_CenterX.get("1.0", END).replace('\n', '')
			_y = self.Str_CenterY.get("1.0", END).replace('\n', '')
			_w = self.Str_Weight.get("1.0", END).replace('\n', '')
			_h = self.Str_Height.get("1.0", END).replace('\n', '')

			self.Treeview.item(self.Focused_Item, text="", values=(_x, _y, _w, _h))
			self.Focused_Item = None
			self.Btn_Update_Area.configure(state=DISABLED)
		

	def Btn_OCR_Preview_Areas(self):
		if self.OCR_File_Path != None:
			if os.path.isfile(self.OCR_File_Path[0]):
				im = cv2.imread(self.OCR_File_Path[0])

				(_h, _w) = im.shape[:2]
				print(_h, _w)
				
				ratio = 720 / _h

				print(ratio)
				if ratio != 1:
					width = int(im.shape[1] * ratio)
					height = int(im.shape[0] * ratio)
					dim = (width, height)
					print('dim', dim)
					im = cv2.resize(im, dim, interpolation = cv2.INTER_AREA)

				for row in self.Treeview.get_children():
					child = self.Treeview.item(row)
					im = cv2.rectangle(im, (child["values"][0], child["values"][1]), (child["values"][0] + child["values"][2], child["values"][1] + child["values"][3]), (255,0,0), 2)

				cv2.imshow("Image", im)
				cv2.waitKey(0)
			else:
				self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])		
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])	

	def Open_OCR_Result_Folder(self):

		try:
			path = self.Function_Correct_Path(os.path.dirname( self.OCR_File_Path[0]))
			_cmd = 'explorer ' + "\"" + str(path) + "\""
			
			subprocess.Popen(_cmd)
		except AttributeError:
			self.Show_Error_Message('Please select source folder.')
			return


	def Btn_OCR_Execute(self):
		'''
		Execute main function
		'''
		Text_Files = self.OCR_File_Path
		Text_Folder =  os.path.dirname( self.OCR_File_Path[0])
		#_temp_text_files = os.listdir(Text_Folder)
		#Text_Files = []
		#for file in _temp_text_files:
		#	file_path = Text_Folder + '/' + file
		#	if os.path.isfile(file_path):
		#		baseName = os.path.basename(file_path)
		#		sourcename, ext = os.path.splitext(baseName)
		#		if 'xls' in ext:
		#			Text_Files.append(file_path)

		match_type_index = self.Match_Type.get()
		if match_type_index == 1:
			exact_match = True
		else:
			exact_match = False

		Db_File = self.BadWord_DB_Path

		Sheet_Name = "Data"
		
		try:
			Sheet_Name = self.BadWord_Data_Sheet_Name.get("1.0", END).replace('\n', '')
		except Exception as e:
			ErrorMsg = ('Error message: ' + str(e))
			print(ErrorMsg)

		Index_Col = "String"
		try:
			Index_Col = self.BadWord_ColumnID.get("1.0", END).replace('\n', '')
		except Exception as e:
			ErrorMsg = ('Error message: ' + str(e))
			print(ErrorMsg)
		

		try:
			self.Background_Color
		except:
			self.Background_Color = 'ffff00'	
		if self.Background_Color == False or self.Background_Color == None:
			self.Background_Color = 'ffff00'
		#print('Background_Color: ', self.Background_Color)
		
		try:
			self.Font_Color
		except:
			self.Font_Color = 'FF0000'	
		if self.Font_Color == False or self.Font_Color == None:
			self.Font_Color = 'FF0000'
		#print('Font_Color: ', self.Font_Color)

		timestamp = "" #Function_Get_TimeStamp()			
		Output_Result_Folder = Text_Folder + '/' + 'Bad_Word_Result_' + str(timestamp)
		if not os.path.isdir(Output_Result_Folder):
			os.mkdir(Output_Result_Folder)
			
		self.BadWord_Check_Process = Process(target=Function_BadWord_Execute, args=(self.Status_Queue, self.Process_Queue, Text_Files, Db_File, Output_Result_Folder, Sheet_Name, Index_Col, exact_match, self.Background_Color, self.Font_Color,))
		self.BadWord_Check_Process.start()
		self.after(DELAY1, self.Wait_For_BadWord_Process)	

	def Wait_For_BadWord_Process(self):
		if (self.BadWord_Check_Process.is_alive()):
			
			try:
				percent = self.Process_Queue.get(0)
				self.BadWord_Progressbar["value"] = percent
				self.BadWord_Progressbar.update()
				#self.Progress.set("Progress: " + str(percent/10) + '%')
			except queue.Empty:
				pass	
			
			try:
				Status = self.Status_Queue.get(0)
				if Status != None:
					self.Write_Debug(Status)
					self.Debugger.insert("end", "\n\r")
					self.Debugger.insert("end", Status)
					self.Debugger.yview(END)
			except queue.Empty:
				pass	
			self.after(DELAY1, self.Wait_For_BadWord_Process)
		else:
			try:
				percent = self.Process_Queue.get(0)
				self.BadWord_Progressbar["value"] = percent
				self.BadWord_Progressbar.update()
				#self.Progress.set("Progress: " + str(percent/10) + '%')
			except queue.Empty:
				pass
			try:
				Status = self.Status_Queue.get(0)
				if Status != None:	
					self.Write_Debug('Bad word check is completed')
					#print(Status)
					self.Debugger.insert("end", "\n\r")
					self.Debugger.insert("end", Status)
					self.Debugger.yview(END)
			except queue.Empty:
				pass
			self.BadWord_Check_Process.terminate()

###########################################################################################
# Process function
###########################################################################################

def Function_BadWord_Execute():

	return

def Function_AutoTest():

	return	

###########################################################################################
# Main loop
###########################################################################################

def main():
	Process_Queue = Queue()
	Result_Queue = Queue()
	Status_Queue = Queue()
	Debug_Queue = Queue()
	
	MyManager = Manager()
	Default_Manager = MyManager.list()
	
	root = Tk()
	My_Queue = {}
	My_Queue['Process_Queue'] = Process_Queue
	My_Queue['Result_Queue'] = Result_Queue
	My_Queue['Status_Queue'] = Status_Queue
	My_Queue['Debug_Queue'] = Debug_Queue

	My_Manager = {}
	My_Manager['Default_Manager'] = Default_Manager

	OCR_Project(root, Queue = My_Queue, Manager = My_Manager,)
	root.mainloop()  


if __name__ == '__main__':
	if sys.platform.startswith('win'):
		multiprocessing.freeze_support()

	main()
