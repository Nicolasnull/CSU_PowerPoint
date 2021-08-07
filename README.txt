#Introduction

Version 1

Hello user!

I hope this program helps you as you create PowerPoints for CSU!
If the program does not work on your computer you may need to 
download python or install the packet python-pptx.

If you don't have python go to https://www.python.org/downloads/
If you don't have python pptx, run installLibraries.bat or 
go to command line and type 'pip install python-pptx' 

Some things you need to know:
Version 1 only contains access to the KJV bible for verses
If this program becomes useful to many people I may download 
more versions. All of the books of the Bible are downloaded
and kept in the KJV folder in this package. If other versions
are downloaded in the same format into another folder, it should 
be easy to add versions in later iterations of this program.

If you need to update info about the club including leaders,
sponsors or meeting times, please edit the 'groupInfo.txt' file
in the appropriate place.

If you would like to pick a particular theme for PowerPoint
go to the 'format.pptx' and choose another design and save.
You can easily change the style after creating too within the PowerPoint software.

You should be able to run this program simply by clicking 'runMe.bat'
This runs the program and then opens up the file once you're done
For the best experience run from the runMe.bat file


Basic function of the program:
- The program I wrote is saved under app.py in this folder. It is in Python and uses the library python-pptx
- If you want to make any changes or add other functionality that would be the place to do it
- To create the title slide it uses the text of the first 3 lines in the groupInfo.txt
- To create the officers slide it uses from the 4th line until the end of the file
- The announcements slide will ask for a list of prayer requests from the user and will loop until
  the user is done entering prayer requests. It will then prompt the user to enter the series the group 
  is currently going through (this is a single line from the user). Finally, the meeting place is entered
  by the user (this is another single line of input)
- The last introductory slide is simply the title of the message and the person presenting it (both are 
  single lines of input)
- The remaining slides are the lesson
- The user will loop through as many lesson slides as needed
- I have coded 3 different options: Bible passages, Header slide and Bullet point slide
- Bible passages will ask the user for the book, chapter, beginning verse and ending verse
  As of right now, the program can do 1 chapter at a time, but you can easily just have back
  to back bible passages with multiple chapters it will just have to be inputted one at a time.
  The passages are being pulled from text files located in the KJV folder. All of those text files
  are formatted the same so the program can easily find the correct passage for the user.
  I did my best to have the verses neatly organized on each slide by limiting the amount of characters 
  per slide. This allows on average 3 verses per slide depending on how big the verses are.
- Header slides are simply a section header with only one line 
  These could be used as questions or whatever you need for a slide with only one line of text
- bullet point slides have a title and then a list of points. simple enough.
  The user will enter a single line for the title and then loop until done with points


I tried to make this user friendly and ran a bunch of tests, but if you find an error please let me know
