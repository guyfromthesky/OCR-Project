#System variable and io handling
import os
import csv
from posixpath import basename
import sys
import random

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

from libs.configmanager import ConfigLoader
from libs.version import get_version
from libs.tkinter_extension import AutocompleteCombobox

import cv2
import numpy as np
import pytesseract

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

		self.Path_Size = 60

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
			
		#Generate UI
		self.Generate_Menu_UI()
		self.Generate_Tab_UI()
		self.init_UI()
		
		self.init_UI_Data()


	# UI init
	def init_UI(self):
	
		self.Generate_OCR_Tool_UI(self.OCR_TOOL)

		self.Generate_OCR_Setting_UI(self.OCR_SETTING)

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
		
		
		self.OCR_SETTING = Frame(self.TAB_CONTROL)
		self.TAB_CONTROL.add(self.OCR_SETTING, text= 'OCR Setting')

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

		
		Radiobutton(Tab, width= 15, text=  self.LanguagePack.Option['Folder'], value=1, variable=self.Browse_Type, command=self.OCR_Setting_Set_Browse_Type).grid(row=Row, column=7,columnspan=2,padx=0, pady=5, sticky=W)
	


		

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

		Radiobutton(Tab, width= 15, text= self.LanguagePack.Option['File'], value=2, variable=self.Browse_Type, command=self.OCR_Setting_Set_Browse_Type).grid(row=Row, column=7,columnspan=2, padx=0, pady=5, sticky=W)
		
		
		Row+=1
		TreeView_Row = 5
		self.Treeview = Treeview(Tab)
		self.Focused_Item = None
		self.Treeview.grid(row=Row, column=1, columnspan=8, rowspan=TreeView_Row, padx=5, pady=5, sticky = N+S+W+E)
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

		verscrlbar.grid(row=Row, column=8, rowspan=TreeView_Row,  sticky = N+S+E)
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
		
		#Row+=1
		#Btn_Preview_Scan = Button(Tab, width = self.Button_Width_Half, text= self.LanguagePack.Button['PreviewArea'], command= self.Btn_OCR_Preview_Scan)
		#Btn_Preview_Scan.grid(row=Row, column=9, columnspan=2,padx=5, pady=5, sticky=W)

		Row+=1

		Btn_Save_Setting = Button(Tab, width = self.Button_Width_Half, text=  self.LanguagePack.Button['SaveConfig'], command= self.Btn_OCR_Save_Config_File)
		Btn_Save_Setting.grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=W)

		Row+=1

		self.Btn_Open_Result = Button(Tab, width = self.Button_Width_Half, text= self.LanguagePack.Button['OpenOutput'], command= self.Open_OCR_Result_Folder)
		self.Btn_Open_Result.grid(row=Row, column=9, columnspan=2,padx=5, pady=5, sticky=W)
		self.Btn_Open_Result.configure(state=DISABLED)

		Row+=1

		Btn_Execute = Button(Tab, width = self.Button_Width_Half, text= "Place Holder", command= None)
		Btn_Execute.grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=W)
		Btn_Execute.configure(state=DISABLED)

		Row+=1
		Label(Tab, text= self.LanguagePack.Label['Debug']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.Debugger = scrolledtext.ScrolledText(Tab, width=110, height=5, undo=False, wrap=WORD, )
		self.Debugger.grid(row=Row, column=2, columnspan=8, padx=5, pady=5, sticky=W+E+N+S)

		Row += 1
		Label(Tab, text= self.LanguagePack.Label['WorkingRes']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
	
		Radiobutton(Tab, width= 10, text=  '720p', value=1, variable=self.Resolution, command= self.OCR_Setting_Set_Working_Resolution).grid(row=Row, column=2, padx=0, pady=5, sticky=W)
		Radiobutton(Tab, width= 10, text=  '1080p', value=2, variable=self.Resolution, command= self.OCR_Setting_Set_Working_Resolution).grid(row=Row, column=3, padx=0, pady=5, sticky=W)
	
	
		GachaAnalyzeEnable = Checkbutton(Tab, text=  self.LanguagePack.Option['GachaScan'], variable = self.GachaAnalyze, command=self.OCR_Setting_Set_Gacha_Analyze)
		GachaAnalyzeEnable.grid(row=Row, column=4,padx=0, pady=5, sticky=W)

		Row += 1
		Label(Tab, text= self.LanguagePack.Label['WorkingLang']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.option_working_language = OptionMenu(Tab, self.WorkingLanguage, *self.language_list, command = self.OCR_Setting_Set_Working_Language)
		self.option_working_language.config(width=self.Button_Width_Full)
		self.option_working_language.grid(row=Row, column=2, padx=5, pady=5, sticky=W)

	


		Row+=1
		Label(Tab, text= self.LanguagePack.Label['Progress']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.progressbar = Progressbar(Tab, orient=HORIZONTAL, length=800,  mode='determinate')
		self.progressbar["maximum"] = 1000
		self.progressbar.grid(row=Row, column=2, columnspan=6, padx=5, pady=5, sticky=E+W)

		Btn_Execute = Button(Tab, width = self.Button_Width_Half, text= self.LanguagePack.Button['Scan'], command= self.Btn_OCR_Execute)
		Btn_Execute.grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=W)

	def Generate_OCR_Setting_UI(self, Tab):
		Row = 1
		Label(Tab, text= self.LanguagePack.Label['TesseractPath']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.Text_TesseractPath = Entry(Tab,width = 100, state="readonly", textvariable=self.TesseractPath)
		self.Text_TesseractPath.grid(row=Row, column=3, columnspan=5, padx=5, pady=5, sticky=E+W)
		Button(Tab, width = self.Button_Width_Full, text=  self.LanguagePack.Button['Browse'], command= self.Btn_Select_Tesseract_Path).grid(row=Row, column=8, columnspan=2, padx=5, pady=5, sticky=E)
		
		Row += 1
		Label(Tab, text= self.LanguagePack.Label['TesseractDataPath']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.Text_TesseractDataPath = Entry(Tab,width = 100, state="readonly", textvariable=self.TesseractDataPath)
		self.Text_TesseractDataPath.grid(row=Row, column=3, columnspan=5, padx=5, pady=5, sticky=E+W)
		Button(Tab, width = self.Button_Width_Full, text=  self.LanguagePack.Button['Browse'], command= self.Btn_Select_Tesseract_Data_Path).grid(row=Row, column=8, columnspan=2, padx=5, pady=5, sticky=E)
		

	

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
		self.TesseractPath = StringVar()
		self.TesseractDataPath = StringVar()
		self.WorkingLanguage = StringVar()

		self.Browse_Type = IntVar()

		self.Resolution = IntVar()
		self.CurrentDataSource = StringVar()


		self.GachaAnalyze = IntVar()

		self.Notice = StringVar()

		self.AppConfig = ConfigLoader()
		self.Configuration = self.AppConfig.Config
		self.AppLanguage  = self.Configuration['OCR_TOOL']['app_lang']
		
		_db_path = self.Configuration['OCR_TOOL']['db_file']
		self.DB_Path.set(_db_path)

		_tesseract_path = self.Configuration['OCR_TOOL']['tess_path']
		pytesseract.pytesseract.tesseract_cmd = _tesseract_path
		self.TesseractPath.set(_tesseract_path)

		_tesseract_data_path = self.Configuration['OCR_TOOL']['tess_data']
		pytesseract.pytesseract.tesseract_cmd = _tesseract_data_path
		self.TesseractDataPath.set(_tesseract_data_path)

		tessdata_dir_config = '--tessdata-dir ' + "\"" + _tesseract_data_path + "\""
		pytesseract.pytesseract.tesseract_cmd = _tesseract_path
		try:
			self.language_list = pytesseract.get_languages(config=tessdata_dir_config)

		except Exception as e:
			print('Error when getting language:', e)
			self.language_list = ['']

		_browse_type = self.Configuration['OCR_TOOL']['browsetype']
		self.Browse_Type.set(_browse_type)

		_resolution = self.Configuration['OCR_TOOL']['resolution']
		self.Resolution.set(_resolution)

		_gacha_scan = self.Configuration['OCR_TOOL']['gachascan']
		self.GachaAnalyze.set(_gacha_scan)
		
	def init_UI_Data(self):
		_working_language = self.Configuration['OCR_TOOL']['scan_lang']
		print('Working language:', _working_language)
		self.WorkingLanguage.set(_working_language)


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

	def Function_Get_TimeStamp(self):		
		now = datetime.now()
		timestamp = str(int(datetime.timestamp(now)))			
		return timestamp

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
		
		self.Btn_Open_Result.configure(state=DISABLED)

		config_file = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'], filetypes = (("Config files", "*.csv *.xlsx"), ), multiple = False)	
		
		if os.path.isfile(config_file):
			print('config_file', config_file)
			self.Str_OCR_Config_Path.set(config_file)
			self.remove_treeview()
			with open(config_file, newline='', encoding='utf-8-sig') as csvfile:
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
		
		self.Btn_Open_Result.configure(state=DISABLED)
		
		_select_type = self.Browse_Type.get()
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
					if ext in ['.jpg','.jpeg','.png']:
						self.OCR_File_Path.append(file_path)

			self.Str_OCR_Image_Path.set(str(folder_name) + '/*')

			self.Write_Debug(self.LanguagePack.ToolTips['DataSelected'] + ": " + str(len(self.OCR_File_Path)))
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])

	def Btn_OCR_Browse_Image_Files(self):
			
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'], filetypes = (("Image files", "*.jpg *.jpeg *png"), ), multiple = True)	
		if filename != "":
			self.OCR_File_Path = list(filename)
			self.Str_OCR_Image_Path.set(str(self.OCR_File_Path[0]))
			
			self.Write_Debug(self.LanguagePack.ToolTips['DataSelected'] + ": " + str(len(self.OCR_File_Path)))
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		
	def Btn_OCR_Select_Area(self):
		
		if self.OCR_File_Path != None:
			_index = random.randint(0, len(self.OCR_File_Path)-1)
			if os.path.isfile(self.OCR_File_Path[_index]):
				im = cv2.imread(self.OCR_File_Path[_index])
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
		if _x == '': _x = 0
		_y = self.Str_CenterY.get("1.0", END).replace('\n', '')
		if _y == '': _y = 0
		_w = self.Str_Weight.get("1.0", END).replace('\n', '')
		if _w == '': _w = 0
		_h = self.Str_Height.get("1.0", END).replace('\n', '')
		if _h == '': _h = 0
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


	def Btn_OCR_Update_Area(self):

		if self.Focused_Item != None:
			_x = self.Str_CenterX.get("1.0", END).replace('\n', '')
			if _x == '': _x = 0
			_y = self.Str_CenterY.get("1.0", END).replace('\n', '')
			if _y == '': _y = 0
			_w = self.Str_Weight.get("1.0", END).replace('\n', '')
			if _w == '': _w = 0
			_h = self.Str_Height.get("1.0", END).replace('\n', '')
			if _h == '': _h = 0
			self.Treeview.item(self.Focused_Item, text="", values=(_x, _y, _w, _h))
			self.Focused_Item = None
			self.Btn_Update_Area.configure(state=DISABLED)
		
	def Function_Load_Img(self, path):
		img = cv2.imread(path)
		(_h, _w) = img.shape[:2]
		_working_res = self.Resolution.get()
		if _working_res == 1:
			_ratio = 720
		else:
			_ratio = 1080

		ratio =  _ratio / _h
		if ratio != 1:
			width = int(img.shape[1] * ratio)
			height = int(img.shape[0] * ratio)
			dim = (width, height)
			img = cv2.resize(img, dim, interpolation = cv2.INTER_AREA)
		
		return img

	def Btn_OCR_Preview_Areas(self):
		if self.OCR_File_Path != None:
			_index = random.randint(0, len(self.OCR_File_Path)-1)
			if os.path.isfile(self.OCR_File_Path[_index]):
				im = self.Function_Load_Img(self.OCR_File_Path[_index])
				for row in self.Treeview.get_children():
					child = self.Treeview.item(row)
					im = cv2.rectangle(im, (child["values"][0], child["values"][1]), (child["values"][0] + child["values"][2], child["values"][1] + child["values"][3]), (255,0,0), 2)

				cv2.imshow("Image", im)
				cv2.waitKey(0)
			else:
				self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])		
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])	

	def Btn_OCR_Preview_Scan(self):
		'''
		Execute main function
		'''
		self.preview_index = random.randint(0, len(self.OCR_File_Path)-1)
		if os.path.isfile(self.OCR_File_Path[self.preview_index]):
			img_file = self.OCR_File_Path[self.preview_index]
			_working_res = self.Resolution.get()
			if _working_res == 1:
				_ratio = 720
			else:
				_ratio = 1080
			
			self._scan_areas = []
			for row in self.Treeview.get_children():
				child = self.Treeview.item(row)
				self._scan_areas.append(child['values'])

			_tess_data = self.TesseractDataPath.get()
			_tess_path = self.TesseractPath.get()
			
			_tess_language = self.WorkingLanguage.get()

			self.BadWord_Check_Process = Process(target=Function_Preview_Scan, args=(self.Result_Queue, self.Process_Queue, _tess_path,_tess_language, _tess_data, img_file, _ratio, self._scan_areas, ))
			
			self.BadWord_Check_Process.start()
			self.after(DELAY1, self.Wait_For_Preview_Process)	

	def Wait_For_Preview_Process(self):
		if (self.BadWord_Check_Process.is_alive()):
			
			try:
				percent = self.Process_Queue.get(0)
				self.progressbar["value"] = percent
				self.progressbar.update()
				#self.Progress.set("Progress: " + str(percent/10) + '%')
			except queue.Empty:
				pass	
			self.after(DELAY1, self.Wait_For_Preview_Process)
		else:
			try:
				percent = self.Process_Queue.get(0)
				self.progressbar["value"] = percent
				self.progressbar.update()
				#self.Progress.set("Progress: " + str(percent/10) + '%')
			except queue.Empty:
				pass
			try:
				Result = self.Result_Queue.get(0)
				if Result != None:	
					index = 0
					img = self.Function_Load_Img(self.OCR_File_Path[self.preview_index])
					
					font = cv2.FONT_HERSHEY_SIMPLEX
					fontScale= 1
					fontColor= (255,255,255)
					lineType= 2

					for area in self._scan_areas:
						_area = (area[0],area[0]+area[2])
						_text = Result[index]
						print(_text)
						img = cv2.rectangle(img, (area[0], area[1]), (area[0] + area[2], area[1] + area[3]), (255,0,0), 2)

						img = cv2.putText(img,_text,_area,font,fontScale,fontColor,lineType)
						index+=1
				cv2.imshow("Image", img)
				cv2.waitKey(0)
			except queue.Empty:
				pass
			self.BadWord_Check_Process.terminate()


	def Open_OCR_Result_Folder(self):

		try:
			path = self.Function_Correct_Path(self.Output_Result_Folder)
			_cmd = 'explorer ' + "\"" + str(path) + "\""
			
			subprocess.Popen(_cmd)
		except AttributeError:
			self.Show_Error_Message('Please select source folder.')
			return


	def Btn_OCR_Execute(self):
		'''
		Execute main function
		'''
		Image_Files = self.OCR_File_Path
		Image_Folder =  os.path.dirname( self.OCR_File_Path[0])

		timestamp = self.Function_Get_TimeStamp()			
		self.Output_Result_Folder = Image_Folder + '/' + 'Scan_Result_' + str(timestamp)
		if not os.path.isdir(self.Output_Result_Folder):
			os.mkdir(self.Output_Result_Folder)
		output_result_file = self.Output_Result_Folder + '/result.csv'
		_ratio = 720	
		_scan_areas = []
		for row in self.Treeview.get_children():
			child = self.Treeview.item(row)
			_scan_areas.append(child['values'])

		_tess_data = self.TesseractDataPath.get()
		_tess_path = self.TesseractPath.get()
 
		_tess_language = self.WorkingLanguage.get()

		self.Btn_Open_Result.configure(state=NORMAL)

		_gacha_analyze = self.GachaAnalyze.get()
		if _gacha_analyze == 1:
			gacha_analyze = True
		else:
			gacha_analyze = False

		self.OCR_Scan_Process = Process(target=Function_Batch_OCR_Execute, args=(self.Result_Queue, self.Status_Queue, self.Process_Queue, _tess_path,_tess_language, _tess_data, Image_Files, output_result_file, _ratio, _scan_areas, gacha_analyze, ))
		
		self.OCR_Scan_Process.start()
		
		self.progressbar["value"] = 0
		self.progressbar.update()

		self.after(DELAY1, self.Wait_For_OCR_Process)

	def Wait_For_OCR_Process(self):
		if (self.OCR_Scan_Process.is_alive()):
			
			try:
				percent = self.Process_Queue.get(0)
				self.progressbar["value"] = percent
				self.progressbar.update()
				#self.Progress.set("Progress: " + str(percent/10) + '%')
			except queue.Empty:
				pass	
			
			try:
				Status = self.Status_Queue.get(0)
				if Status != None:
					self.Write_Debug(Status)
					
			except queue.Empty:
				pass	
			self.after(DELAY1, self.Wait_For_OCR_Process)
		else:
			while True:
				try:
					percent = self.Process_Queue.get(0)
					self.progressbar["value"] = percent
					self.progressbar.update()
				except queue.Empty:
					break
			while True:
				try:
					Status = self.Status_Queue.get(0)
					if Status != None:	
						self.Write_Debug('Bad word check is completed')
						#print(Status)
				except queue.Empty:
					break
			self.OCR_Scan_Process.terminate()
			self.Write_Debug(self.LanguagePack.ToolTips['Completed'])

###########################################################################################
# OCR Setting
###########################################################################################

	def Btn_Select_Tesseract_Path(self):
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectDB'],filetypes = (("Executable files","*.exe" ), ), )	
		if os.path.isfile(filename):
			_tess_path = self.CorrectPath(filename)
			self.AppConfig.Save_Config(self.AppConfig.Ocr_Tool_Config_Path, 'OCR_TOOL', 'tess_path', _tess_path, True)
			pytesseract.pytesseract.tesseract_cmd = _tess_path
			self.TesseractPath.set(_tess_path)
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['TessNotSelect'])

	def Btn_Select_Tesseract_Data_Path(self):
		folder_name = filedialog.askdirectory(title =  self.LanguagePack.ToolTips['SelectSource'],)	
		if os.path.isdir(folder_name):
			folder_name = self.CorrectPath(folder_name)
			self.TesseractDataPath.set(folder_name)

			self.AppConfig.Save_Config(self.AppConfig.Ocr_Tool_Config_Path, 'OCR_TOOL', 'tess_data', folder_name, True)

			self.Write_Debug(self.LanguagePack.ToolTips['DataSelected'] + ": " + folder_name)
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		

	def OCR_Setting_Set_Gacha_Analyze(self):
		_gacha_analyze = self.GachaAnalyze.get()
		if _gacha_analyze == 1:
			_status = 'enabled'
		else:
			_status = 'disabled'
		
		self.AppConfig.Save_Config(self.AppConfig.Ocr_Tool_Config_Path, 'OCR_TOOL', 'gachascan', _gacha_analyze)

		self.Write_Debug(self.LanguagePack.ToolTips['AnalyzeGachaResult'] + str(_status))

	def OCR_Setting_Set_Browse_Type(self):
		_browse_type = self.Browse_Type.get()
		if _browse_type == 1:
			_status = 'folder'
		else:
			_status = 'file'
		
		self.AppConfig.Save_Config(self.AppConfig.Ocr_Tool_Config_Path, 'OCR_TOOL', 'browsetype', _browse_type)

		self.Write_Debug(self.LanguagePack.ToolTips['BrowseTypeUpdate'] + str(_status))

	def OCR_Setting_Set_Working_Resolution(self):
		_resolution_index = self.Resolution.get()
		if _resolution_index == 1:
			self.WorkingResolution = 720
		else:
			self.WorkingResolution = 1080
		
		self.AppConfig.Save_Config(self.AppConfig.Ocr_Tool_Config_Path, 'OCR_TOOL', 'resolution', _resolution_index)

		self.Write_Debug(self.LanguagePack.ToolTips['SetResolution'] + str(self.WorkingResolution) + 'p')

	def OCR_Setting_Set_Working_Language(self, select_value):
		
		self.AppConfig.Save_Config(self.AppConfig.Ocr_Tool_Config_Path, 'OCR_TOOL', 'scan_lang', select_value)
		
		self.Write_Debug(self.LanguagePack.ToolTips['SetScanLanguage'] + str(select_value))
	

###########################################################################################
# Process function - Batch scan
###########################################################################################

def Function_Batch_OCR_Execute(
	Result_Queue, Status_Queue, Process_Queue, tess_path, tess_language, tess_data, image_files, result_file, ratio, scan_areas, gacha_analyze, **kwargs):
	
	advanced_tessdata_dir_config = '--psm 7 --tessdata-dir ' + '"' + tess_data + '"'

	if tess_language == '':
		tess_language = 'kor'
	Status_Queue.put('Scan language: ' + tess_language)

	number_of_processes = multiprocessing.cpu_count()

	_task_list = []
	processes = []
	for image in image_files:
		str_filename = str(image)
		_task_list.append(str_filename)

	if gacha_analyze:
		_total = int(len(_task_list) * 1.1)
	else:
		_total = len(_task_list)	

	_complete = 0
	Area_Name = ['Area_' + str(i) for i in range(len(scan_areas))]
	Title = ['FileName'] + Area_Name

	with open(result_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
		writer = csv.writer(csvfile)
		writer.writerow(Title)
	
	while len(_task_list) > 0:
		if len(_task_list) > number_of_processes:
			_new_task_count = number_of_processes
		else:
			_new_task_count = len(_task_list)

		for w in range(_new_task_count):

			input_file = _task_list[0]

			p = Process(target=Get_Text_From_Single_Image, args=(tess_path, tess_language, advanced_tessdata_dir_config, input_file, ratio, scan_areas, result_file,))

			del _task_list[0]
			processes.append(p)
			p.start()

		for p in processes :
			p.join()
			_complete+=1
		
		percent = ShowProgress(_complete, _total)
		Process_Queue.put(percent)

	if gacha_analyze:
		Function_Analyze_Gacha_Data(result_file, Area_Name)
		percent = ShowProgress(_total, _total)
		Process_Queue.put(percent)

def Get_Text_From_Single_Image(tess_path, tess_language, advanced_tessdata_dir_config, input_image, ratio, scan_areas, result_file,):

	pytesseract.pytesseract.tesseract_cmd = tess_path

	_img = Load_Image_by_Ratio(input_image, ratio)

	_result = []

	for area in scan_areas:
		imCrop = _img[int(area[1]):int(area[1]+area[3]), int(area[0]):int(area[0]+area[2])]
		imCrop = Function_Pre_Processing_Image(imCrop)
		ocr = Get_Text(imCrop, tess_language, advanced_tessdata_dir_config)
		_result.append(ocr)
	#print(_result)
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

def Function_Analyze_Gacha_Data(_raw_data, col_name):
	_output_dir = os.path.dirname(_raw_data)
	analyze_result_file = _output_dir + "/analyze_result.csv"
	_gacha = {}
	
	with open(_raw_data, newline='', encoding='utf-8-sig') as csvfile:
		reader = csv.DictReader(csvfile)
		for row in reader:
			for component in col_name:
				component_name = row[component]
				if component_name in _gacha:
					_gacha[component_name] +=1
				else:
					_gacha[component_name] =1

	with open(analyze_result_file, 'a', newline='', encoding='utf-8-sig') as csvfile:
		fieldnames = ['Components', 'Amount']
		writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
		writer.writeheader()
		for component in _gacha:
			writer.writerow({'Components': component, 'Amount': _gacha[component]})
	print('Analyze done')			

###########################################################################################
# Process function - Preview scan
###########################################################################################

def Function_Preview_Scan(
	Result_Queue, Process_Queue, tess_path, tess_language, test_data, image_files, ratio, scan_areas, **kwargs):
	pytesseract.pytesseract.tesseract_cmd = tess_path
	tessdata_dir_config = '--tessdata-dir ' + '"' + test_data + '"'
	if tess_language == '':
		tess_language = 'kor'
	
	_result = []
	_counter = 0
	_total = len(scan_areas)
	_result = []

	_img = Load_Image_by_Ratio(image_files, ratio)


	for area in scan_areas:
		imCrop = _img[int(area[1]):int(area[1]+area[3]), int(area[0]):int(area[0]+area[2])]
		imCrop = Function_Pre_Processing_Image(imCrop)
		
		ocr = Get_Text(imCrop, tess_language, tessdata_dir_config)
	
		_result.append(ocr)
		_counter+=1	
		Process_Queue.put(int(_counter*1000/_total))

	Result_Queue.put(_result)

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

def Get_Text(img, tess_language, tessdata_dir_config):
	ocr = pytesseract.image_to_string(img, lang = tess_language, config=tessdata_dir_config)
	ocr = ocr.replace("\n", "")
	ocr = ocr.replace("\r", "")  
	ocr = ocr.replace("\x0c", "") 
	return ocr

def Function_Pre_Processing_Image(img):
	img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
	img = cv2.resize(img, None, fx=0.5, fy=0.5, interpolation=cv2.INTER_AREA)
	img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
	img = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_LINEAR)
	#img = cv2.blur(img,(5,5))
	#img = cv2.GaussianBlur(img, (5, 5), 0)
	#img = cv2.medianBlur(img, 3)
	#img = cv2.bilateralFilter(img,9,75,75)
	#cv2.threshold(img,127,255,cv2.THRESH_BINARY)
	img = image_smoothening(img)

	return	img

def image_smoothening(img):
	BINARY_THREHOLD = 100
	ret1, th1 = cv2.threshold(img, BINARY_THREHOLD, 255, cv2.THRESH_BINARY)
	ret2, th2 = cv2.threshold(th1, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
	#blur = cv2.GaussianBlur(th2, (1, 1), 0)
	ret3, th3 = cv2.threshold(th2, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
	return th3


def ShowProgress(Counter, TotalProcess):
	#os.system('CLS') 
	percent = int(1000 * Counter / TotalProcess)
	#print("Current progress: " +  str(Counter) + '/ ' + str(TotalProcess))
	return percent

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
