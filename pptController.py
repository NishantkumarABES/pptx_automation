import win32com.client
import time

app = win32com.client.Dispatch("PowerPoint.Application")
presentation = app.Presentations.Open(FileName=r"D:\User Data\Desktop\nishant\Mini Project\2-year\presentation\MiniPrj.pptx", ReadOnly=1)
presentation.SlideShowSettings.run()


def moveRight():
    time.sleep(1)
    presentation.SlideShowWindow.View.Next()
def moveLeft():
    time.sleep(1)
    presentation.SlideShowWindow.View.Previous()
    
def Exit():
    presentation.SlideShowWindow.View.Exit()
    app.Quit()


# Importing Libraries
import io
import cv2
import time
import pygame
import pyaudio
import imutils
import numpy as np
from gtts import gTTS
import mediapipe as mp 
import win32com.client



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

CHUNK = 1024  
RATE = 44100  
THRESHOLD = 1500  

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
    app = win32com.client.Dispatch("PowerPoint.Application")
    presentation = app.Presentations.Open(FileName=r"D:\User Data\Desktop\nishant\Mini Project\2-year\presentation\MiniPrj.pptx", ReadOnly=1)
    presentation.SlideShowSettings.run()
    def moveRight():
        time.sleep(1)
        presentation.SlideShowWindow.View.Next()
    def moveLeft():
        time.sleep(1)
        presentation.SlideShowWindow.View.Previous()
    def Exit():
        presentation.SlideShowWindow.View.Exit()
        app.Quit()

    from google.protobuf.json_format import MessageToDict 
    mpHands = mp.solutions.hands 
    hands = mpHands.Hands( 
        static_image_mode=False, 
        model_complexity=1, 
        min_detection_confidence=0.75, 
        min_tracking_confidence=0.75, 
        max_num_hands=2)


    cap = cv2.VideoCapture(0) 

    while True:  
        success, img = cap.read()  
        imgRGB = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        results = hands.process(imgRGB)  
        if results.multi_hand_landmarks: 
                if len(results.multi_handedness) == 2: 
                        cv2.putText(img, 'Both Hands', (250, 50),cv2.FONT_HERSHEY_COMPLEX, 0.9,(0, 255, 0), 2)
                        print("Exit")
                        speak("Program Terminated")
                        Exit()
                        break
                else: 
                    for i in results.multi_handedness: 
                        label = MessageToDict(i)['classification'][0]['label'] 
                            if label == 'Left': 
                                cv2.putText(img, label+' Hand', (20, 50),cv2.FONT_HERSHEY_COMPLEX, 0.9, (0, 255, 0), 2)
                                moveLeft()
                            if label == 'Right':  
                                cv2.putText(img, label+' Hand', (460, 50), cv2.FONT_HERSHEY_COMPLEX,0.9, (0, 255, 0), 2)
                                moveRight()
 
 
        if cv2.waitKey(1) & 0xff == ord('q'): 
                break

