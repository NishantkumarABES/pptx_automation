import io,cv2
import os
import time, pygame
import pyaudio, imutils
import numpy as np
from gtts import gTTS
import mediapipe as mp 
import win32com.client

CHUNK = 1024  
RATE = 44100  
THRESHOLD = 1500

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QHBoxLayout

pptx_app = win32com.client.Dispatch("PowerPoint.Application")

from google.protobuf.json_format import MessageToDict 
mpHands = mp.solutions.hands 
hands = mpHands.Hands( 
        static_image_mode=False, 
        model_complexity=1, 
        min_detection_confidence=0.75, 
        min_tracking_confidence=0.75, 
        max_num_hands=2)

def speak(text):
    print("\n" + text)
    sound = gTTS(text=text, lang='hi')
    audio_data = io.BytesIO()
    sound.write_to_fp(audio_data)
    audio_data.seek(0)
    pygame.mixer.init()
    pygame.mixer.music.load(audio_data)
    pygame.mixer.music.play()
    while pygame.mixer.music.get_busy():
        continue

def activation_function():
    p = pyaudio.PyAudio()
    stream = p.open(format=pyaudio.paInt16, channels=1, rate=RATE, input=True, frames_per_buffer=CHUNK)
    print("Listening for Activation...")
    active = False
    while True:
        data = stream.read(CHUNK)
        audio_data = np.frombuffer(data, dtype=np.int16)
        audio_data_abs = np.abs(audio_data)
        avg_energy = np.mean(audio_data_abs)
        if avg_energy > THRESHOLD:
            print("Finger snap detected!")
            speak("System Activated!")
            active = True
            break

    stream.stop_stream()
    stream.close()  
    p.terminate()

    
def pptx_controller(file_path):
    presentation = pptx_app.Presentations.Open(FileName=file_path, ReadOnly=1)
    presentation.SlideShowSettings.run()
    def moveRight():
        time.sleep(1)
        presentation.SlideShowWindow.View.Next()
    def moveLeft():
        time.sleep(1)
        presentation.SlideShowWindow.View.Previous()
    def Exit():
        presentation.SlideShowWindow.View.Exit()
        pptx_app.Quit()

    cap = cv2.VideoCapture(0) 

    while True:  
        success, img = cap.read()  
        imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        results = hands.process(imgRGB)  
        if results.multi_hand_landmarks: 
            for i in results.multi_handedness: 
                label = MessageToDict(i)['classification'][0]['label'] 
                if label == 'Left': moveLeft()
                if label == 'Right': moveRight()

        if cv2.waitKey(1) & 0xff == ord('q'): 
                break


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(324, 321)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.file_path = None
        
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(0, 0, 321, 191))
        self.label.setPixmap(QtGui.QPixmap("PowerPoint Automation.png"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")
        
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(10, 230, 111, 23))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.upload_file)
        
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(130, 230, 75, 23))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.start)
        self.pushButton_2.setEnabled(False)
        
        self.message_label = QtWidgets.QLabel(self.centralwidget)
        self.message_label.setGeometry(QtCore.QRect(10, 270, 191, 16))
        self.message_label.setObjectName("message_label")
        #self.message_label.setText("")
        
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 324, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.pushButton.setText(_translate("MainWindow", "Upload presentation"))
        self.pushButton_2.setText(_translate("MainWindow", "Start"))

    def upload_file(self):
        self.file_path = QFileDialog.getOpenFileName()
        if self.file_path:
            self.pushButton_2.setEnabled(True)
            self.message_label.setText('File uploaded successfully!')
        else:
            self.message_label.setText('No file selected.')

    def start(self):
        self.message_label.setText('System is ready')
        time.sleep(2)
        activation_function()
        relative_path = "\\\\".join(self.file_path[0].split("/"))
    
        try: pptx_controller(relative_path)
        except Exception as E: print(E)
        

            
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
