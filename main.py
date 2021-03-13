import win32com.client
from os.path import isfile, join
import os
from tkinter.filedialog import askdirectory
from alive_progress import alive_bar
import time

powerpoint_directory = askdirectory(title="Select Folder source of PPT")
PptApp = win32com.client.Dispatch("Powerpoint.Application")
PptApp.Visible = True
corpus = [str(f) for f in os.listdir(powerpoint_directory) if not f.startswith('.') and isfile(join(powerpoint_directory, f))]

with alive_bar(len(corpus)) as bar:
    for filename in corpus:
        path = powerpoint_directory + "/" + filename
        print(path)
        PPtPresentation = PptApp.Presentations.Open(os.path.abspath(path))
        output_name = filename.split(".")[0]
        print("Output name: " + output_name)
        save_as_name = powerpoint_directory + "/output/" + output_name
        PPtPresentation.SaveAs(os.path.abspath(save_as_name), 24)
        PPtPresentation.close()
        bar()

PptApp.Quit()
print("Done!")
