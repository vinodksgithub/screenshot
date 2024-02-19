# screenshot
This application allows you to capture screenshots of your screen and add comments to them, creating Word  documents with the captured images and your notes


How to run ?
python capture_part_of_screen.py

Detais
The provided code creates a graphical user interface (GUI) application in Python using the tkinter library.
 This application allows you to capture screenshots of your screen and add comments to them, creating Word
 documents with the captured images and your notes. Here's a breakdown of the functionalities:

Main features:

Capture entire screen: Takes a screenshot of the whole screen and saves it as a Word document with an optional comment you can add.
Trim screenshot: Allows you to specify the height and width in millimeters to trim the captured image before adding it to the document.
Custom capture: Lets you select a specific area of the screen by capturing two sets of coordinates and saves it as a custom screenshot in a Word document with comments (optional).
Choose save location: You can set a preferred folder to save the generated Word documents.
Timer option: Provides a 10-second timer to automatically capture after choosing the option.
Mouse coordinate tracking: Displays the current X and Y coordinates of your mouse cursor on the screen, helpful for selecting custom capture areas.
Technical details:

The code uses tkinter to create the user interface with buttons, labels, text entry fields, and checkboxes.
PIL (Python Imaging Library) is used for image acquisition and manipulation (screenshot capture, cropping, saving).
docx library is used for creating and adding images to Word documents.
threading module enables running timer and capture actions simultaneously.
pyautogui is used for retrieving the user's mouse cursor position on the screen.
Overall, this code offers a comprehensive tool for capturing screenshots, customizing them, and saving them with comments in Word documents, potentially useful for documentation, bug reporting, or personal note-taking with visual references.
