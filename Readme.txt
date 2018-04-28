How To Run?

Open the folder where the directory is saved. If Python is installed on your machine, proceed further, else install it and then proceed.
1.    Open cmd and go the project folder.
2.    Run the command : python questionStatus@SPOJ.py 

What does this Script Do?

It accepts command line arguments and stores them as usernames for SPOJ. Then it open each user's profile and visit each question on user's profile, subsequently each question's submission data like Submission Id, Submission time, Solution run time, Language used is stored. Now each users data is being written to an excel sheet made on the fly and stored in a folder named - worksheets.

Modules Used: sys, webbrowser, requests, xlsxwriter
