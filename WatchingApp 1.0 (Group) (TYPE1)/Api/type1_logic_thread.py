import sys
from PyQt5.QtCore import QThread, pyqtSignal
from csv import writer
import csv

###progress bar
import time
from datetime import datetime
from pytz import timezone

import env as ENV


class External(QThread):
    """
    Runs a time_cnter thread.
    """
    countChanged = pyqtSignal(int)
    printChanged = pyqtSignal(str)

    last_print = "C"
    first_flag = 1
    opposite = 0
    files = []
    start_hour = 0
    end_hour = 24

    # requirement
    Enter_B_TVA_need = 0
    Enter_S_TVA_need = 0
    Print_C_IN_B_TVA = 0
    Print_C_IN_S_TVA = 0

    #initial the data before start new cycle
    def initialData(self) :
        print ("=========================================")
        print ("File names : ",self.files)
        print ("Printing B TVA value : ",self.Enter_B_TVA_need)
        print ("Printing S TVA value : ",self.Enter_S_TVA_need)
        print ("Printing C during B TVA value : ",self.Print_C_IN_B_TVA)
        print ("Printing C during S TVA value : ",self.Print_C_IN_S_TVA)
        print ("Start : ",self.start_hour)
        print ("End : ",self.end_hour)
        print ("OPPOSITE : ", self.opposite)
        print ("=========================================")
        return

    # return current GMT+3 time
    def getCurrentGmtTime(self) :
        # Current time in GMT+3
        now_gmt = datetime.now(timezone('GMT'))
        return now_gmt.astimezone(timezone('Europe/Moscow'))

    # monitor the cycle time monitora and return the remain time
    def getRemainTime_monitorTVA(self) -> float:
        #determine whether TVA is on right now
        cur_time = self.getCurrentGmtTime()
        minute = cur_time.minute
        hour = cur_time.hour
        second = cur_time.second

        if minute == 0 and second <= 5 and hour <= self.end_hour and hour >= self.start_hour:
            self.monitorTVA()

        cur_time = self.getCurrentGmtTime()
        minute = cur_time.minute

        #remain the next time
        return minute * 60 + cur_time.second

    def log_detail(self, detailData) :
        file_name = ENV.MASTERFILEPATH + 'Type1' + '_info.csv'
        #represent the writed data into the console
        print(detailData)

        while(1):
            try :
                with open(file_name, 'a+', newline='') as write_obj:
                    # Create a writer object from csv module
                    csv_writer = writer(write_obj)
                    # Add contents of list as last row in the csv file
                    csv_writer.writerow(detailData)
                return
            except IOError as x:
                print  ("couldn't read this file on this spot")

    # remove one last row for printing some thing like (BB, SS)
    def remove_last_row(self) :
        file = open(ENV.MASTERFILEPATH + ENV.TYPE1_PRINT, "r+")
        lines = file.readlines()
        lines.pop()
        file = open(ENV.MASTERFILEPATH + ENV.TYPE1_PRINT, "w+")
        file.writelines(lines)

    #print B on the log and csv and write the log into the info
    def printBSC(self, print_char, cur_TVAs, last_TVAs) :

        if self.last_print == print_char :
            return

        str_data = " "
        format = "%Y.%m.%d,%H:%M:%S"
        # Current time in GMT
        now_gmt = datetime.now(timezone('GMT'))
        gmt_time = now_gmt.astimezone(timezone('Europe/Moscow'))
        date_str = gmt_time.strftime(format)

        for filename in self.files :
            printData = [
                filename[0:filename.find("_")] + 
                ','+ 'Date:'+ date_str +
                ',Type1' + ':'+ print_char
            ]

            logData = [
                filename[0:filename.find("_")] + 
                ','+ 'Date:'+ date_str +
                ',Type3' + ':'+ print_char + " " +
                "Current TVAs : " + str(cur_TVAs) + " "
                "Previous TVAs : " + str(last_TVAs)
            ]

            # Appending a row to csv with missing entries
            if print_char != "DO NOTHING":
                self.last_print = print_char
                self.append_list_as_row(printData)

            self.printChanged.emit(str_data.join(printData))
            self.log_detail(logData)

        #Print any signs if current group any opposite
        if self.opposite != 0 :
            if print_char == "B" :
                print_char = "S"
            elif print_char == "S" :
                print_char = "B"

            printData = [
                self.opposite[0:self.opposite.find("_")] + 
                ','+ 'Date:'+ date_str+
                ',Type3' + ':'+ print_char
            ]

            logData = [
                self.opposite[0:self.opposite.find("_")] + 
                ','+ 'Date:'+ date_str +
                ',Type3' + ':'+ print_char + "  " + 
                "Current TVAs : " + str(cur_TVAs) + " "
                "Previous TVAs : " + str(last_TVAs)
            ]

            # Appending a row to csv with missing entries
            if print_char != "DO NOTHING":
                self.last_print = print_char
                self.append_list_as_row(printData)

            self.printChanged.emit(str_data.join(printData))
            self.log_detail(logData)

    #write the new data on the last row of the data
    def append_list_as_row(self, data):
        # Open file in append mode
        file_name = ENV.MASTERFILEPATH + 'Type3' + '.csv'

        #represent the writed data into the console
        print(data)

        while(1):
            try :
                with open(file_name, 'a+', newline='') as write_obj:
                    # Create a writer object from csv module
                    csv_writer = writer(write_obj)
                    # Add contents of list as last row in the csv file
                    csv_writer.writerow(data)
                return
            except IOError as x:
                print  ("couldn't read this file on this spot")


    #read the previous row from the csv file
    def read_CUR_PREVIOUS_row (self, filename) :
        last_row = []
        previous_row = []
        last_3_row = []
        filename = ENV.FILEPATH + 'IndexValuePanelData_' + filename + ".csv"
        while(1):
            try :
                with open(filename, mode='r') as csv_file:
                    csv_reader = csv.reader(csv_file)
                    for row in csv_reader :
                        last_3_row = previous_row
                        previous_row = last_row
                        last_row = row
                if previous_row[2] == "Index 1" :
                    return [-1, 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0]

                data = {}
                data['last_row'] = previous_row
                data['previous_row'] = last_3_row

                return data
            except IOError as x:
                print  ("couldn't read this file on this spot")

    #read the last row from the csv file
    def read_last_row(self, filename) :
        last_row = []

        while(1):
            try :
                with open(filename, mode='r') as csv_file:
                    csv_reader = csv.reader(csv_file)
                    line_cnt = 0
                    for row in csv_reader :
                        last_row = row
                return last_row
            except IOError as x:
                print  ("couldn't read this file on this spot")


    # get the TVA infos from the last row of the source asset
    def monitorTVA(self) :

        self.first_flag = 0

        zero_flag = 0
        TVA_SUM_B = 0
        TVA_SUM_S = 0
        cur_TVAs = []
        previous_TVAs = []
        TVA_cnt = len(self.files) * 4

        for filename in self.files :
            data = self.read_CUR_PREVIOUS_row(filename)
            last_row = data['last_row']
            previous_row = data['previous_row']

            #change date values to 0 because of floating converting.
            last_row[0] = 0
            last_row[15] = 0
            previous_row[0] = 0
            previous_row[15] = 0

            #make an array which reflects the current TVA values and close price
            cur_row = []
            for item in last_row :
                cur_row.append(float(item))

            previous_values = []
            for item in previous_row :
                previous_values.append(float(item))


            cur_TVAs.append(cur_row[16])
            cur_TVAs.append(cur_row[17])
            cur_TVAs.append(cur_row[18])
            cur_TVAs.append(cur_row[19])

            previous_TVAs.append(previous_values[16])
            previous_TVAs.append(previous_values[17])
            previous_TVAs.append(previous_values[18])
            previous_TVAs.append(previous_values[19])

            # check any 0 values in previous and current TVA values
            if 0 in cur_TVAs or 0 in previous_TVAs:
                zero_flag = 1

            if cur_TVAs[0] < previous_TVAs[0] :
                TVA_SUM_S += 1
            elif cur_TVAs[0] > previous_TVAs[0] :
                TVA_SUM_B += 1

            if cur_TVAs[1] < previous_TVAs[1] :
                TVA_SUM_S += 1
            elif cur_TVAs[1] > previous_TVAs[1] :
                TVA_SUM_B += 1

            if cur_TVAs[2] < previous_TVAs[2] :
                TVA_SUM_S += 1
            elif cur_TVAs[2] > previous_TVAs[2] :
                TVA_SUM_B += 1

            if cur_TVAs[3] < previous_TVAs[3] :
                TVA_SUM_S += 1
            elif cur_TVAs[3] > previous_TVAs[3] :
                TVA_SUM_B += 1

        print ("Here we go for printing anything")
        print ("================================")
        print ("Count For B :  " , TVA_SUM_B)
        print ("Count For S :  " , TVA_SUM_S)
        print ("Count For 0 :  " , TVA_cnt - TVA_SUM_B - TVA_SUM_S)
        print ("================================")

        if self.last_print == "C":
            #Print Do nothing if there is any zero in current or previous TVA values
            if zero_flag == 1 :
                self.printBSC("DO NOTHING", cur_TVAs, previous_TVAs)
                return

            #Printing B LOGIC
            if TVA_SUM_B >= self.Enter_B_TVA_need:
                self.printBSC("B", cur_TVAs, previous_TVAs)
                return
            #Printing S LOGIC
            if TVA_SUM_S >= self.Enter_S_TVA_need:
                self.printBSC("S", cur_TVAs, previous_TVAs)
                return

        #Triggering C during B TVA logic
        elif self.last_print == "B":
            #Print Do nothing if there is any zero in current or previous TVA values
            if zero_flag == 1 :
                self.printBSC("C", cur_TVAs, previous_TVAs)
                return

            if TVA_SUM_B <= self.Print_C_IN_B_TVA :
                self.printBSC("C", cur_TVAs, previous_TVAs)
                return

        #Triggering C during S TVA logic
        elif self.last_print == "S" :
            #Print Do nothing if there is any zero in current or previous TVA values
            if zero_flag == 1 :
                self.printBSC("C", cur_TVAs, previous_TVAs)
                return
            if TVA_SUM_S <= self.Print_C_IN_S_TVA :
                self.printBSC("C", cur_TVAs, previous_TVAs)
                return

        self.printBSC("DO NOTHING", cur_TVAs, previous_TVAs)

    def monitorIndexValues(self) :
        #check the time is on avaliable

        cur_time = self.getCurrentGmtTime()
        hour = cur_time.hour

        if hour < self.start_hour or hour > self.end_hour :
            return
        if hour == self.end_hour and cur_time.minute >= 1 :
            return

    def run(self):
        #initial datas
        self.initialData()

        #inital remaing time and monitor TVA with the rule
        time_cnt = self.getRemainTime_monitorTVA()
        while time_cnt < ENV.CYCLE_TIME:
            # get the start time perf_count
            start_time = time.perf_counter()

            self.countChanged.emit(time_cnt)

            # if self.first_flag == 0 and time_cnt % ENV.CHECK_INDEX_DURATION_TIME < ENV.REFRESH_TIME :
            #     self.monitorIndexValues()

            # get the end time perf_count
            end_time = time.perf_counter()
            time.sleep(ENV.REFRESH_TIME - (end_time - start_time))
            time_cnt += ENV.REFRESH_TIME

        self.run()
