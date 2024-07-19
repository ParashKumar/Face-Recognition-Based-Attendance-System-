import cv2
import face_recognition
import os
from datetime import datetime
import openpyxl

# Path to the directory containing images of known faces
known_faces_path = 'known_faces'

# Load known faces and corresponding names
known_face_encodings = []
known_face_names = []

for file_name in os.listdir(known_faces_path):
    if file_name.endswith('.jpg') or file_name.endswith('.png'):
        image_path = os.path.join(known_faces_path, file_name)
        img = face_recognition.load_image_file(image_path)
        face_encoding = face_recognition.face_encodings(img)[0]
        known_face_encodings.append(face_encoding)
        known_face_names.append(os.path.splitext(file_name)[0])

# Create or load Excel file for attendance tracking
excel_file_path = 'Attendance.xlsx'

# Check if the Excel file exists, and create it if not
if not os.path.isfile(excel_file_path):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet['A1'] = 'Name'
    sheet['B1'] = 'Date Time'
    wb.save(excel_file_path)

# Start webcam capture
cap = cv2.VideoCapture(0)

while True:
    # Capture frame-by-frame
    ret, frame = cap.read()

    # Find all face locations and face encodings in the current frame
    face_locations = face_recognition.face_locations(frame)
    face_encodings = face_recognition.face_encodings(frame, face_locations)

    for face_encoding, (top, right, bottom, left) in zip(face_encodings, face_locations):
        # Check if the face matches any known faces
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)

        name = 'Unknown'

        if True in matches:
            first_match_index = matches.index(True)
            name = known_face_names[first_match_index]

            # Write attendance to Excel file
            workbook = openpyxl.load_workbook(excel_file_path)
            sheet = workbook.active
            row = (name, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            sheet.append(row)
            workbook.save(excel_file_path)
            workbook.close()

        # Draw rectangle and label around the face
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
        cv2.putText(frame, name, (left, top - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.75, (0, 255, 0), 2)

    # Display the resulting frame
    cv2.imshow('Video', frame)

    # Break the loop when 'q' key is pressed
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Release the webcam and close all windows
cap.release()
cv2.destroyAllWindows()
