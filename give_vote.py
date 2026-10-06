from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(str1)

video = cv2.VideoCapture(0)
facedetect = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
if not os.path.exists('data/'):
    os.makedirs('data/')

with open('data/names.pkl', 'rb') as f:
    LABELS = pickle.load(f)

with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)
imgBackground = cv2.imread("background.png")

COL_NAMES = ['NAME', 'VOTE', 'DATE', 'TIME']

def check_if_exists(value):
    try:
        with open("Votes.csv", "r") as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if row and row[0] == value:
                    return True
    except FileNotFoundError:
        print("File not found or unable to open the CSV file.")
    return False

# --- define container coordinates (adjust to match your blank box) ---
container_x, container_y = 300, 300  # top-left corner of blank box
container_w, container_h = 640, 480  # width and height of blank box

# --- fullscreen window setup ---
cv2.namedWindow("SMART VOTING SYSTEM", cv2.WND_PROP_FULLSCREEN)
cv2.setWindowProperty("SMART VOTING SYSTEM", cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)

while True:
    ret, frame = video.read()
    if not ret:
        break

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)
    
    output = None
    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w]
        resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
        output = knn.predict(resized_img)
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M-%S")
        exist = os.path.isfile("Votes.csv")
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 0, 255), 1)
        cv2.putText(frame, str(output[0]), (x, y-15), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)
        attendance = [output[0], timestamp]
    
    # --- proportional resize and centering inside container ---
    h, w = frame.shape[:2]
    scale = min(container_w / w, container_h / h)  # keep aspect ratio
    new_w, new_h = int(w * scale), int(h * scale)
    frame_resized = cv2.resize(frame, (new_w, new_h))

    # create a blank container area
    container_area = np.zeros((container_h, container_w, 3), dtype=np.uint8)

    # center the resized frame inside the container
    y_offset = (container_h - new_h) // 2
    x_offset = (container_w - new_w) // 2
    container_area[y_offset:y_offset+new_h, x_offset:x_offset+new_w] = frame_resized

    # paste into background
    imgBackground[container_y:container_y+container_h, container_x:container_x+container_w] = container_area

    cv2.imshow("SMART VOTING SYSTEM", imgBackground)
    k = cv2.waitKey(1)
    
    if output is not None:
        voter_exist = check_if_exists(output[0])
        if voter_exist:
            speak("you are already counted thankyou")
            break

        if k == ord('1'):
            speak("YOUR VOTE HAS BEEN RECORDED")
            time.sleep(5)
            with open("Votes.csv", "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                if not exist:
                    writer.writerow(COL_NAMES)
                attendance = [output[0], "BJP", date, timestamp]
                writer.writerow(attendance)
            speak("THANK YOU FOR PARTICIPATING IN THE ELECTIONS")
            break

        if k == ord('2'):
            speak("YOUR VOTE HAS BEEN RECORDED")
            time.sleep(5)
            with open("Votes.csv", "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                if not exist:
                    writer.writerow(COL_NAMES)
                attendance = [output[0], "CONGRESS", date, timestamp]
                writer.writerow(attendance)
            speak("Thank you for making your voice heard—every vote counts")
            break

        if k == ord('3'):
            speak("YOUR VOTE HAS BEEN RECORDED")
            time.sleep(5)
            with open("Votes.csv", "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                if not exist:
                    writer.writerow(COL_NAMES)
                attendance = [output[0], "AAP", date, timestamp]
                writer.writerow(attendance)
            speak("Grateful for your participation in shaping our future. Your vote truly matters")
            break

        if k == ord('4'):
            speak("YOUR VOTE HAS BEEN RECORDED")
            time.sleep(5)
            with open("Votes.csv", "a", newline="") as csvfile:
                writer = csv.writer(csvfile)
                if not exist:
                    writer.writerow(COL_NAMES)
                attendance = [output[0], "NOTA", date, timestamp]
                writer.writerow(attendance)
            speak("THANKS FOR STANDING UP AND BEING COUNTED. DEMOCRACY THRIVES BECAUSE OF VOTERS LIKE YOU.")
            break

video.release()
cv2.destroyAllWindows()