import pyautogui, webbrowser
import os
import time

import pyperclip, subprocess

#have a text file called pyautoguidemo.txt
#in the text file are five lines: first line: John, second line: Smith, third line: john@smith.com, fourth line: 1231231234, fifth line: Hello
#this program will fill out a form with the info in the text file (rather than the data being hard coded in the python program)

"""os.startfile("Notepad.exe") #open notepad
time.sleep(1)
pyautogui.press(["alt", "f", "o"]) #open new file
time.sleep(1)
pyautogui.typewrite(os.getcwd()+ "\\" + "pyautoguidemo.txt")
time.sleep(1)
pyautogui.write("enter")"""

#Reduce 7 lines of code (above) to 1 line of code (below) using Popen() function
subprocess.Popen(['C:\\Windows\\notepad.exe', 'C:\\Users\\mmcclain\\python\\pyautoguidemo.txt'])

pyautogui.PAUSE = 0.5
print('Ensure that the browser window is active and the form is loaded!')
webbrowser.open('https://www.meghanmcclain.com', new=2)
pyautogui.click(-1150,39)
time.sleep(1)
pyautogui.scroll(-50000)

print('>>> 5-SECOND PAUSE TO LET USER PRESS CTRL-C <<<')
time.sleep(1)

#copy line 1 in text file, paste in first field
#copy line 2 in text file, paste in second field, and so on
myfile = open("pyautoguidemo.txt", "r")
myline = myfile.readline()
count = 1
while myline:
    print(myline)
    pyperclip.copy(myline)
    pyautogui.click(-1522,-2)
    pyautogui.press('\t', presses=count)
    count += 1
    pyautogui.hotkey('ctrl', 'v')
    myline = myfile.readline()
myfile.close()
