# Introduction

This program was written by Nicolas Null and Jalen Wayt. The conceptual idea is credited to Nicolas Null.
It was designed to automate the process of creating PowerPoint presentations for CSU.

# Instructions
### 1) Installing Python and required Libraries

If you do not already have Python installed, you can install it at [Python Download](https://www.python.org/downloads/).
<br /><br />
Currently, the only library used is known as python-pptx. In order to install the library, a batch file is included for your convenience. If you would prefer to use the command line, you can type "pip install python-pptx".

### 2) Updating Organization Information 

The program pulls information from the groupInfo.txt file. If any changes are made to the leadership of CSU, please make those changes there.

### 3) Changing Format of PowerPoints 

The program uses the format.pptx file to determine the theme to use. If you open the file, you can select a different theme and/or design. This will be incorporated into the generated presentation.

### 4) Function of the Program

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

### 5) Additional Information 

The first version of this program only contains the verses for the King James Version (KJV). We do intend on adding some other languages upon request from potential users of this program.
In order to run the program, you can execute the 'runme.bat' file.
