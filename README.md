# EETesttool
EETesttool is basically helps manual testers for easily capturing screenshots into excel workbooks. This helps to speed up manual testing and avoid copy paste task, just a button click does the magic for you :)

This tool is developed using java and Jframe is used for UI
One the application is started a small Jframe with filecontrols will be on the desktop and will be there till you close it.
There are two jtextfields with labels as workbook and worksheet
and the last control is capture button
Whenever you click on capture, the tool captures active window screenshot and copies it into given work book in given worksheet
You can keep on clikcing as many screenshots you want and it will be saved in given workbook, there will be four rows space between two screenshots in same worksheet.
If given file contains contents, scripting append the screenshots and preserve the existing data from workbook.
Worksheets can be named after test cases, whener you change name of the worksheet and if its not existing worksheet in given workbook, script will create new worksheet.This way you will be able to seperate out screenshots test casewise and after complete execution few labeling changes will be sufficient to convert this file test evidence.
Steps to set up: Download source code, build executable jar, create a .bat file to execute this executable jar. Double click on .bat file. A java framewindow will apear, give th path of you test evidence workbook in workbook test box and test case name in worksheet testbox. Start your test case execution, wherever you need screenshot click on capture button, crop the image as per requiremnt and move with next step of your execution. Cropped screenshot is automatically saved to you evidence file in given worksheet. With new test case change name of worksheet and repeat above steps.
