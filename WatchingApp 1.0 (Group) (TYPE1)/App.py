from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton,QProgressBar, QDialog, QListWidget,QGridLayout
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QTimer, QTime, Qt, pyqtSlot, QThread, pyqtSignal

import pandas as pd

import sys, os, traceback, types
import pathlib
from pytz import timezone

#import timezone module
from datetime import datetime
import pytz 
#import enviroment values
import env as ENV
#import thread apis
import Api.type1_logic_thread as TYPE1_LOGIC_THREAD

#When the Estimate Button is clicked
class MainWindow(QWidget):
	"""
	Simple dialog that consists of a Progress Bar and a Button.
	Clicking on the button results in the start of a timer and
	updates the progress bar.
	"""
	def __init__(self):
		QWidget.__init__(self)
		self.setStyleSheet("color: black; background-color: rgb(0,162,232);")
		self.setGeometry(60,80,500,600)

		self.setWindowTitle('Watching SVC Files FOR TYPE1-ABC')
		
		#button
		self.button = QPushButton('Start', self)
		self.button.setStyleSheet("color: black; font-size : 23px; border : 3px solid black")
		self.button.setGeometry(40, 30, 130, 40)
		self.button.clicked.connect(self.onButtonClick)

		#progressbar
		self.progress = QProgressBar(self)
		self.progress.setGeometry(40, 100, 420, 15)
		self.progress.setMaximum(ENV.CYCLE_TIME)


		#List label
		list_label = QLabel(self)
		list_label.setStyleSheet("color: black; font-size : 23px;")
		list_label.setText("PRINTING LOG")
		list_label.setGeometry(60, 160, 200, 30)
		#listbox
		self.listwidget = QListWidget(self)
		self.listwidget.setStyleSheet("color: black; border : 3px solid black; font-size : 16px")
		self.listwidget.setGeometry(40, 200, 420, 300)

		self.time_label = QLabel(self)
		self.time_label.setStyleSheet("color: black; font-size : 20px; border : 3px solid black")
		self.time_label.setText("")
		self.time_label.setGeometry(40, 520, 300, 30)

		timer = QTimer(self)
		timer.timeout.connect(self.showTime)
		timer.start(1000)

		self.show()

	def onButtonClick(self):

		try:
		    with open(ENV.MASTERFILEPATH + ENV.TYPE1_PRINT):
		        os.remove(ENV.MASTERFILEPATH + ENV.TYPE1_PRINT)
		        print ("Old TYPE1.csv file will be removed and create a new one.")
		except IOError:
			print ("New TYPE1.csv file will be created")

		try:
		    with open(ENV.MASTERFILEPATH + "Type1_info.csv"):
		        os.remove(ENV.MASTERFILEPATH + "Type1_info.csv")
		        print ("Old TYPE1_info.csv file will be removed and create a new one.")
		except IOError:
			print ("New TYPE1_info.csv file will be created")

		#for type1 processes
		print ("======================================OPENING MASTERSHEET EXCEL FILES================================")
		print ("-------------------------------------- TYPE1 MASTERSHEET ------------------------------------")
		
		#read the master sheet file for each csv files		
		mastersheet = pd.read_excel (ENV.MASTERFILEPATH + ENV.TYPE1_MASTER_FILENAME)
		records = mastersheet.to_records()

		groupmastersheet = pd.read_excel(ENV.MASTERFILEPATH + ENV.GROUP_TYPE1_MASTER_FILENAME)
		group_cnt = len(groupmastersheet.columns)
		key_values = list(groupmastersheet.keys())
		if key_values[-1] == "SKIP n/a" :
			group_cnt -= 1

		self.type1_Run = []

		for x in range(0,group_cnt) :
			groupmastersheet = pd.read_excel(ENV.MASTERFILEPATH + ENV.GROUP_TYPE1_MASTER_FILENAME, usecols = [x])
			group_temp = groupmastersheet.to_records()

			#run new process for new group
			run_index = len(self.type1_Run)
			self.type1_Run.append(TYPE1_LOGIC_THREAD.External())
			self.type1_Run[run_index].files = []
			index = 0
			while 1 :
				if group_temp[index][1].find("B: ") != -1 :
					break
				#new file object created
				#read file name
				self.type1_Run[run_index].files.append(group_temp[index][1])
				index += 1


			#for triggering B of TVA condition
			self.type1_Run[run_index].Enter_B_TVA_need = int(group_temp[index][1].replace("B: ", "").replace(" >", ""))
			index += 1

			#for triggering S of TVA condition
			self.type1_Run[run_index].Enter_S_TVA_need = int(group_temp[index][1].replace("S: < ", ""))
			index += 3


			#for printing C in B of TVA condition
			self.type1_Run[run_index].Print_C_IN_B_TVA = int(group_temp[index][1].replace("C for B: <=", ""))
			index += 1

			#for printing C in S of TVA condition
			self.type1_Run[run_index].Print_C_IN_S_TVA = int(group_temp[index][1].replace("C for S: =>", ""))
			index += 2


			#for starting time for this group
			self.type1_Run[run_index].start_hour = int(group_temp[index][1].replace("Start : ", ""))
			index += 1

			#for ending time for this group
			self.type1_Run[run_index].end_hour = int(group_temp[index][1].replace("End : ", ""))
			index += 2

			if len(group_temp) > index and group_temp[index][1] == "OPPOSITE" :
				self.type1_Run[run_index].opposite = group_temp[index + 1][1]

			self.type1_Run[run_index].countChanged.connect(self.onCountChanged)
			self.type1_Run[run_index].printChanged.connect(self.onPrintChanged)
			
			self.type1_Run[run_index].start()


	def onCountChanged(self, value):
		self.progress.setValue(value)

	def onPrintChanged(self, value) :
		self.listwidget.addItems([value])

	def showTime(self) :
		format = "%H:%M:%S"
		# Current time in UTC
		now_gmt = datetime.now(timezone('GMT'))
		gmt_time = now_gmt.astimezone(timezone('Europe/Moscow'))
		self.time_label.setText("GMT+3      " + gmt_time.strftime(format))


if __name__ == "__main__":
	app = QApplication(sys.argv)
	screen = MainWindow()
	screen.show()
	sys.exit(app.exec_())
