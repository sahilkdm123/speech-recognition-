import os
import speech_recognition as sr
import tkinter as tk
from tkinter import messagebox, Toplevel
from tkcalendar import Calendar
import pandas as pd
from datetime import datetime

def recognize_speech():
    r = sr.Recognizer()
    numbers = []
    with sr.Microphone() as source:
        print("Speak:")
        try:
            audio = r.listen(source, timeout=5, phrase_time_limit=20)
            text = r.recognize_google(audio)
            print(f"Recognized text: {text}")  # Print the recognized text
            numbers = text.split()
        except sr.WaitTimeoutError:
            print("No speech detected")
        except sr.UnknownValueError:
            print("Could not understand audio")
        except sr.RequestError as e:
            print("Could not request results; {0}".format(e))
    return numbers

def verify_roll_no(roll_nos):
    def print_sel():
        return cal.selection_get()

    root = tk.Toplevel()
    MsgBox = tk.messagebox.askquestion ('Cross Verification','Are these your roll numbers: ' + ', '.join(roll_nos) + '?', icon='warning')
    if MsgBox == 'yes':
        top = Toplevel(root)
        cal = Calendar(top, font="Arial 14", selectmode='day', cursor="hand1", year=2022, month=2, day=5)
        cal.pack(fill="both", expand=True)
        tk.Button(top, text="OK", command=root.destroy).pack()  # Close the dialog box when the button is clicked
        root.mainloop()
        return print_sel()
    else:
        root.destroy()  # Close the Tkinter window if the roll numbers are not verified
        return False

def update_sheet(roll_nos, date):
    # Check if the file exists
    if os.path.isfile('roll_numbers.xlsx'):
        # If the file exists, read it into a DataFrame
        df = pd.read_excel('roll_numbers.xlsx')
    else:
        # If the file doesn't exist, create a new DataFrame with roll numbers from 1 to 101
        df = pd.DataFrame({'Roll Number': range(1, 101)})

    # If the date column doesn't exist, create it and fill with 'A'
    if date not in df.columns:
        df[date] = 'A'

    # For each roll number in the image, put 'P' in the corresponding cell
    for roll_no in roll_nos:
        df.loc[df['Roll Number'] == int(roll_no), date] = 'P'

    # Convert all the column labels to string type, except for 'Roll Number'
    df.columns = [str(col) if col != 'Roll Number' else col for col in df.columns]

    # Sort the DataFrame by column names (dates), excluding 'Roll Number'
    df = df.set_index('Roll Number').sort_index(axis=1).reset_index()

    # Write the DataFrame to an Excel file
    df.to_excel('roll_numbers.xlsx', index=False)

# Main program
roll_nos = recognize_speech() 
print(f"Recognized roll numbers: {roll_nos}")  # Print the recognized roll numbers
date = verify_roll_no(roll_nos)
print(f"Selected date: {date}")  # Print the selected date
if date:
    update_sheet(roll_nos, date)