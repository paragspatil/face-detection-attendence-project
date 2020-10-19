from PyQt5.QtGui import QStandardItemModel
from PyQt5.QtWidgets import QApplication, QPushButton, QMainWindow, QWidget, QDialog, QGroupBox, QHBoxLayout, \
    QVBoxLayout, \
    QLabel, QTableWidgetItem, QTableWidget, QHeaderView, QComboBox
from PyQt5 import QtGui
from PyQt5.QtCore import QRect
import os
import cv2
import numpy as np
import face_recognition
import time
from PyQt5.QtGui import QIcon, QPixmap
from ui_main_window import *
from PyQt5.QtGui import QImage
from datetime import datetime
import pandas as pd
import sys
import xlsxwriter


class Window(QDialog):
    def __init__(self):
        super().__init__()

        self.isAttendence = False
        self.title = "Face recognition Layout"
        self.left = 700
        self.top = 700
        self.width = 1000
        self.height = 1000
        self.IconName = "resorces/face-recognition.png"
        self.initWindow()

    def initWindow(self):
        self.setWindowIcon(QtGui.QIcon(self.IconName))
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.SelectClassLayout()
        self.UtilityactionsWindow()
        self.createTable()
        vbox = QVBoxLayout()
        vbox.addWidget(self.groupbox)
        vbox.addWidget(self.groupbox2)
        vbox.addWidget(self.tableWidget)
        self.cameraGroupBox = QGroupBox()
        self.cameraGroupBox.setMinimumHeight(600)
        self.initCameraBox()
        vbox.addWidget(self.cameraGroupBox)
        self.setLayout(vbox)
        self.show()

    def SelectClassLayout(self):
        self.groupbox = QGroupBox("selecting class")
        hboxlayout = QHBoxLayout()

        label = QLabel("Select class")
        label.setMinimumHeight(40)
        hboxlayout.addWidget(label)

        self.comboBox = QComboBox(self)
        self.comboBox.setToolTip("select class")
        classes = os.listdir("classes")
        numOfClasses = len(classes)
        for c in classes:
            self.comboBox.addItem(c)

        self.comboBox.setMinimumHeight(40)

        hboxlayout.addWidget(self.comboBox)
        self.groupbox.setLayout(hboxlayout)

    def UtilityactionsWindow(self):
        self.groupbox2 = QGroupBox("utility actions")
        hboxlayout = QHBoxLayout()

        addstudentButton = QPushButton("add new student")
        addstudentButton.setToolTip("click here to add new student to selected class")
        addstudentButton.setMinimumHeight(40)
        hboxlayout.addWidget(addstudentButton)

        self.startattendenceButton = QPushButton("Start Attendance session")
        self.startattendenceButton.setToolTip("Start Attendance session for selected class")
        self.startattendenceButton.setMinimumHeight(40)
        self.startattendenceButton.clicked.connect(self.StartAttendenceSession)
        hboxlayout.addWidget(self.startattendenceButton)

        self.groupbox2.setLayout(hboxlayout)

    def createTable(self):
        self.tableWidget = QTableWidget()

        # Row count
        self.tableWidget.setRowCount(80)

        # Column count
        self.tableWidget.setColumnCount(4)
        """
        self.tableWidget.setItem(0, 0, QTableWidgetItem("Name"))
        self.tableWidget.setItem(0, 1, QTableWidgetItem("City"))
        self.tableWidget.setItem(1, 0, QTableWidgetItem("Aloysius"))
        self.tableWidget.setItem(1, 1, QTableWidgetItem("Indore"))
        self.tableWidget.setItem(2, 0, QTableWidgetItem("Alan"))
        self.tableWidget.setItem(2, 1, QTableWidgetItem("Bhopal"))
        self.tableWidget.setItem(3, 0, QTableWidgetItem("Arnavi"))
        self.tableWidget.setItem(3, 1, QTableWidgetItem("Mandsaur"))
        """
        self.tableWidget.setHorizontalHeaderLabels(["Roll No", "Name", "Attendece Status", "Time Recorded"])

        # Table will fit the screen horizontally
        self.tableWidget.horizontalHeader().setStretchLastSection(True)
        self.tableWidget.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)

    def initCameraBox(self):
        camerahboxlayout = QVBoxLayout()
        self.exportToExcelButton = QPushButton("Export to excel")
        self.exportToExcelButton.setMinimumHeight(40)
        self.exportToExcelButton.clicked.connect(self.Exporttoexcel)
        camerahboxlayout.addWidget(self.exportToExcelButton)

        self.cameraoutput = QLabel()
        self.cameraoutput.setMinimumHeight(200)
        camerahboxlayout.addWidget(self.cameraoutput)

        self.cameraGroupBox.setLayout(camerahboxlayout)

    def StartAttendenceSession(self):
        if not self.isAttendence:
            self.isAttendence = True
            self.startattendenceButton.setText("stop Attendance session")
            print("clicked")
            studentImages = []
            studentImagesEncodings = []
            self.listofstudentRollnos = []
            self.attendanceStatus = []
            self.timeRecorded = []
            path = "classes/" + self.comboBox.currentText() + "/Students Data/"
            self.listOfstudents = os.listdir(path)
            i = 0
            for student in self.listOfstudents:
                print(path + student + "/" + student + ".jpg")
                currentImage = cv2.imread(path + student + "/" + student + ".jpg")
                studentImages.append(currentImage)
                self.tableWidget.setItem(i, 1, QTableWidgetItem(student))
                self.tableWidget.setItem(i, 0, QTableWidgetItem(i))
                self.tableWidget.setItem(i, 2, QTableWidgetItem("Absent"))
                self.listofstudentRollnos.append(i + 1)
                self.attendanceStatus.append("Absent")
                self.timeRecorded.append("not recorded")
                i = i + 1

            for img in studentImages:
                img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
                encode = face_recognition.face_encodings(img)[0]
                studentImagesEncodings.append(encode)

            self.cap = cv2.VideoCapture(0)

            while self.isAttendence:
                success, img = self.cap.read()

                imgs = cv2.resize(img, (0, 0), None, 0.25, 0.25)
                imgs = cv2.cvtColor(imgs, cv2.COLOR_BGR2RGB)

                facecurrentframe = face_recognition.face_locations(imgs)
                # print(facecurrentframe)
                print(facecurrentframe)
                encodecurrentframe = face_recognition.face_encodings(imgs, facecurrentframe)
                # y1, x2, y2, x1 = facecurrentframe
                # y1, x2, y2, x1 = y1 * 4, x2 * 4, y2 * 4, x1 * 4
                # cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), 2)
                # cv2.rectangle(img, (x1, y2 - 35), (x2, y2), (0, 255, 0), cv2.FILLED)
                # cv2.putText(img, listOfstudents[matchIndex], (x1 + 6, y2 - 6), cv2.FONT_HERSHEY_COMPLEX, 1, (0, 255, 0), 2)
                self.displayImage(img, 1)
                cv2.waitKey()

                for encodeface, faceLoc in zip(encodecurrentframe, facecurrentframe):
                    matches = face_recognition.compare_faces(studentImagesEncodings, encodeface)
                    faceDis = face_recognition.face_distance(studentImagesEncodings, encodeface)
                    matchIndex = np.argmin(faceDis)

                    if matches[matchIndex]:
                        print(self.listOfstudents[matchIndex])
                        self.tableWidget.setItem(matchIndex, 2, QTableWidgetItem("Present"))
                        now = datetime.now()
                        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
                        self.tableWidget.setItem(matchIndex, 3, QTableWidgetItem(dt_string))
                        self.attendanceStatus[matchIndex] = "Present"
                        self.timeRecorded[matchIndex] = dt_string

        else:
            self.isAttendence = False
            self.cap.release()
            self.cameraoutput.clear()
            self.startattendenceButton.setText("Start Attendance session")

    def Exporttoexcel(self):
        i = 0
        if len(self.listOfstudents) > 0:
            now = datetime.now()
            dt_string = now.strftime("%d-%m-%Y %H-%M-%S")
            print(dt_string)
            print(os.path.isdir("classes/" + self.comboBox.currentText() + "/attendence recods"))
            # writing to excel
            workbook = xlsxwriter.Workbook(
                "classes/" + self.comboBox.currentText() + "/attendence recods/" + dt_string + '.xlsx')

            # By default worksheet names in the spreadsheet will be
            # Sheet1, Sheet2 etc., but we can also specify a name.
            worksheet = workbook.add_worksheet("My sheet")

            # Some data we want to write to the worksheet.

            # Start from the first cell. Rows and
            # columns are zero indexed.
            worksheet.write(0, 0, "Roll No")
            worksheet.write(0, 1, "Names")
            worksheet.write(0, 2, "Attendance Status")
            worksheet.write(0, 3, "Time Recorded")
            row = 1
            col = 0

            # Iterate over the data and write it out row by row.
            i = 0
            print("here")
            while i < len(self.listOfstudents):
                worksheet.write(i + 1, 0, self.listofstudentRollnos[i])
                worksheet.write(i + 1, 1, self.listOfstudents[i])
                worksheet.write(i + 1, 2, self.attendanceStatus[i])
                worksheet.write(i + 1, 3, self.timeRecorded[i])
                i += 1

            workbook.close()

    def displayImage(self, img, window=1):
        qformat = QImage.Format_Indexed8

        if len(img.shape) == 3:
            if img.shape[2] == 4:
                qformat = QImage.Format_RGBA8888

            else:
                qformat = QImage.Format_RGB888

        img = QImage(img, img.shape[1], img.shape[0], qformat)
        img = img.rgbSwapped()
        self.cameraoutput.setPixmap(QPixmap.fromImage(img))
        self.cameraoutput.setAlignment(QtCore.Qt.AlignCenter | QtCore.Qt.AlignVCenter)


if __name__ == "__main__":
    App = QApplication(sys.argv)
    window = Window()
    sys.exit(App.exec_())
