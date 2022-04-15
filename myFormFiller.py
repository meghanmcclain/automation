#open browser to my website and fill out my contact form using the following info

import pyautogui, time, webbrowser

formData = [{'fname': 'Linda', 'lname': 'McClain', 'email': 'example@example.com',
            'phone': '(123)-123-1234', 'message': 'Tell Bob I said hi.'},
            {'fname': 'Kevin', 'lname': 'McClain', 'email': 'example@example.com', 'phone': '(123)-123-1234',
            'message': 'n/a'},
            {'fname': 'Bridget', 'lname': 'McClain', 'email': 'example@example.com',
            'phone': '(123)-123-1234', 'message': 'Please take the puppets out of the break room.'},
            {'fname': 'Meredith', 'lname': 'McClain', 'email': 'example@example.com',
            'phone': '(123)-123-1234', 'message': 'Protect the innocent. Serve the public trust. Uphold the law.'},
           ] #list of 4 dictionaries

submitAnotherLink = (-1145, 101)

pyautogui.PAUSE = 0.5
print('Ensure that the browser window is active and the form is loaded!')
webbrowser.open('https://www.meghanmcclain.com', new=2) #open in new tab
pyautogui.click(-1150,39)
#time.sleep(5)
pyautogui.scroll(-50000) #scroll all the way down
#time.sleep(5)

for person in formData:
    # Give the user a chance to kill the script.
    print('>>> 5-SECOND PAUSE TO LET USER PRESS CTRL-C <<<')
    time.sleep(1)
    pyautogui.click(-1150,39)
    #time.sleep(5)

    # Wait until the form page has loaded.
    print('Entering %s info...' % (person['fname']))
    pyautogui.write(['\t'])

    # Fill out the First Name Field.
    pyautogui.write(person['fname'] + '\t')

    # Fill out the Last Name field.
    pyautogui.write(person['lname'] + '\t')

    # Fill out the Email Field.
    pyautogui.write(person['email'] + '\t')

    # Fill out the Phone Number field.
    pyautogui.write(person['phone'] + '\t')

    # Fill out the Message field.
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.write(person['message'] + '\t')

    # Click Submit.
    time.sleep(0.5) # Wait for the button to activate.
    pyautogui.press('enter')

    # Wait until form page has loaded.
    print('Submitted form.')
    time.sleep(1)

    # Click the Submit another response link.
    pyautogui.click(submitAnotherLink[0], submitAnotherLink[1])
