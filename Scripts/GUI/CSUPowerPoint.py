from pptx import Presentation
from pptx.util import Pt
from tkinter import *

#Defined as the functions needed to make a proper powerpoint for our meetings
class CSUPowerPoint:
    def __init__(self, tkinterInterface):
        #Set the Tkinter Interface Variable
        self.interface = tkinterInterface
        self.interface.title("CSU PowerPoint Generator")
        #The presentation uses the format/design of the chosen powerpoint
        self.presentation = Presentation("../../Files/format.pptx")
        #Has all officer information
        self.groupInfoFile = open("../../Files/groupInfo.txt", "r")
        
        #Start with introduction page
        self.titleSlide()
        self.officersSlide()
        self.introduction()

    #Function to simplify the calls to create a slide, add it to the presentation, and assign text to it
    def setSlideInfo(self, layout, titleText):
        #The layout of the slide (title slide, information slide, etc)
        slideLayout = self.presentation.slide_layouts[layout]
        #Add the slide in the specified layout
        slide = self.presentation.slides.add_slide(slideLayout)
        #Create a textbox and add the text specified
        textBox = slide.shapes.title
        textBox.text = titleText
        
        #Returns the slide layout, created slide and the created text box
        return (slideLayout, slide, textBox)
    
    #Gets the input until the user enters something
    def getUserInput(self):
        userInput = input("> ")
        while(len(userInput) == 0):
            userInput = input("> ")

        return userInput
    
    #Gets the chapter and verse from the user
    def getChapterAndVerses(self, printStr):
        print(str(printStr))
        output = int(input("> "))

        return output
    
    #Returns the passage as a whole
    def getPassages(self, book, chapter, startVerse, endVerse):
        passage = []
        for lines in book:
            if not ":" in lines:
                continue
            #Gets the index of the colon
            firstColon = lines.index(":")
            #Assigns the current chapter number
            currentChapter = int(lines[0:firstColon])
            #Gets the second colon
            secondColon = lines.index(":", firstColon + 1,)
            #Assigns the current verse
            currentVerse = int(lines[firstColon+1:secondColon])
            #If the chapter equals the current chapter of the file and the verses match, then append the information
            if currentChapter == chapter and currentVerse >= startVerse and currentVerse <= endVerse:
                passage.append(lines[firstColon+1:])

        return passage

    #Check the passage for irregularities
    def checkPassage(self, userInput):
        #If the input starts with a digit for the book, then make sure it is set to "1 John" for example
        if userInput[0].isdigit():
            userInput = userInput[0] + " " + userInput[1].upper() + userInput[2:]
        elif userInput == "songofsolomon":#only book with spaces in it within the Bible
            userInput = "Song of Solomon"
        else:#Otherwise, set the first letter to a capital
            userInput = userInput[0].upper() + userInput[1:]

        return userInput

    #Performs the introduction page in the GUI
    def introduction(self):
        #Create the Introductory Frame
        self.introductoryFrame = Frame(self.interface)
        self.introductoryFrame.pack()
        #Creates an Introductory Message
        introductionLabel = Label(self.introductoryFrame, text="Welcome to the CSU PowerPoint Generator V1!\nClick the \"Next\" Button to continue!")
        introductionLabel.pack()
        #Adds a next button to move to the next page
        nextButton = Button(self.introductoryFrame, text="Next", fg="green", command=self.announcementsSlide)
        nextButton.pack()
        #Adds an exit button to close the application
        exitButton = Button(self.introductoryFrame, text="Exit", fg="red", command=self.interface.destroy)
        exitButton.pack()

    #Creates the title slide for the meetings
    def titleSlide(self):
        #Sets the slide layout, slide and textbox
        (slideLayout, slide, textBox) = self.setSlideInfo(0, self.groupInfoFile.readline())

        #Assigns the placeholder for the text box to assign more information to it
        textBox = slide.placeholders[1]
        textBox.text = self.groupInfoFile.readline() + "\n" + self.groupInfoFile.readline()

    #Displays all of the information of the officers of the organization in the groupInfoFile.txt
    def officersSlide(self):
        #Sets the slide layout, slide and textbox
        (slideLayout, slide, textBox) = self.setSlideInfo(1, self.groupInfoFile.readline())

        #Assigns a shape to the textbox
        textBox = slide.shapes[1]
        #Assigns a frame to the textbox
        textFrame = textBox.text_frame
        #Adds a paragraph
        lines = textFrame.paragraphs[0]
        #Reads the first line of the file
        lines.text = self.groupInfoFile.readline()
        lines.font.size = Pt(14)
        
        #Iterates through the file until it reaches the end
        person = self.groupInfoFile.readline()
        i=1
        while person != "":
            #Adds a paragraph for each person
            lines = textFrame.add_paragraph()
            #This is their information (Title: Name)
            lines.text = person
            lines.font.size = Pt(14)
            person = self.groupInfoFile.readline()
            i += 1

    def addPrayers(self, prayerFocus, textFrame):
        prayerFocus = prayerFocus.splitlines()
        for index in range(len(prayerFocus)):
            if(prayerFocus[index].startswith("  ")):
                lines = textFrame.add_paragraph()
                lines.text = prayerFocus[index]
                lines.level = 2
            else:
                lines = textFrame.add_paragraph()
                lines.text = prayerFocus[index]
                lines.level = 1


    def savePresentation(self):
        self.presentation.save("your_powerpoint.pptx")
        self.interface.destroy()

    #Creates the announcements slide
    def announcementsSlide(self):
        #Destroys all elements created within the introductory frame
        #The title and officers slides are generated based on the file input
        self.introductoryFrame.destroy()
        #Creates a new frame for the announcements input
        self.announcementsFrame = Frame(self.interface)
        self.announcementsFrame.pack()

        #Sets the slide layout, slide and textbox
        (slideLayout, slide, textBox) = self.setSlideInfo(1, "Announcements")
        #Adds a shape to the textbox
        textBox = slide.shapes[1]
        #Gets the text frame of the textbox
        textFrame = textBox.text_frame
        #Adds text to the frame
        lines = textFrame
        lines.text = "Prayer Focus"
        lines.level = 0
        
        #Prayer Focus (Section to be edited to work with GUI)
        announcementsLabel = Label(self.announcementsFrame, text="Enter prayer focuses below: ")
        announcementsLabel.pack()

        prayerFocusText = Text(self.announcementsFrame)
        prayerFocusText.pack()

        '''
        while True:
            #Adds a paragraph for each prayer focus in the frame
            lines = textFrame.add_paragraph()
            #Gets user input until it equals 'done'
            userInput = self.getUserInput()
            if userInput.lower().replace(" ", "") == "done":
                break
            lines.text = userInput
            lines.level = 1

        #Current Series
        #Adds a paragraph for the series information
        lines = textFrame.add_paragraph()
        print("Enter the name of the series you are going through: ")
        #Gets user input on the series
        userInput = self.getUserInput()
        lines.text = "Series: " + userInput
        
        #Location of Meetings
        #Adds a paragraphs for the location of the meetings
        lines = textFrame.add_paragraph()
        lines = textFrame.add_paragraph()
        print("\nEnter where you are holding meetings: ")
        #Gets user input on the meeting place
        userInput = self.getUserInput()
        lines.text = "Meeting place: " + userInput
        '''
        #Adds a next button to move to the next page
        nextButton = Button(self.announcementsFrame, text="Next", fg="green", command=lambda: self.addPrayers(prayerFocusText.get('1.0', 'end-1c'), textFrame))
        nextButton.pack()
        #Adds an exit button to close the application
        exitButton = Button(self.announcementsFrame, text="Exit", fg="red", command=self.savePresentation)
        exitButton.pack()
    
    #Creates a lesson header slide
    def lessonHeaderSlide(self):
        #Gets the title of the lesson from the user
        print("\nEnter the title of the lesson: ")
        userInput = self.getUserInput()

        #Sets the slide layout, slide and textbox
        slideLayout, slide, textBox = self.setSlideInfo(2, userInput)

        #Gets the name of the person teaching the lesson
        print("\nEnter the person teaching: ")
        userInput = self.getUserInput()
        textBox = slide.placeholders[1]
        textBox.text = userInput
    
    #Creates a bible passage using the KJV folder (Helper Method)
    def biblePassage(self):
        try:
            #Gets the book from th euser
            print("\nEnter book of the Bible: ")
            #Sets it to lowercase after replacing the spaces with empty strings
            userInput = input("> ").replace(" ", "").lower()
            #Grabs the book text file
            book = open('../Files/KJV/' + userInput + '.txt', 'r')
            #Asks the user for the chapter
            chapter = self.getChapterAndVerses("Chapter: ")
            #Asks the user for the starting verse
            startVerse = self.getChapterAndVerses("Starting Verse: ")
            #Asks the user for the ending verse
            endVerse = self.getChapterAndVerses("Ending Verse: ")

            #If the end verse is less than the start verse, raise an error
            if endVerse < startVerse:
                raise(LookupError)
            
            #Iterate through the specified book to find the chapter
            passage = self.getPassages(book, chapter, startVerse, endVerse)

            #If the passage length is 0, raise a lookup error
            if(len(passage) == 0):
                raise(LookupError)
            
            #Check for digits (1 John), or spaces (Song of Solomon)
            userInput = self.checkPassage(userInput)

            #If they are equal, then we only put one verse
            if(startVerse == endVerse):
                title = userInput + " " + str(chapter) + ":" + str(startVerse)
            else:#if they are not, then we have multiple verses denoted as (Book Chapter:StartVerse - Endverse)
                title = userInput + " " + str(chapter) + ":" + str(startVerse) + "-" + passage[len(passage) - 1][:passage[len(passage)-1].index(":")]

            #Iterate through the passage with a maximum of 9 lines per slide
            linesUsed = 9
            for lines in passage:
                #Gets the character in the verse
                charInVerse = len(lines) / 50 + 1
                #If it is greater than 9, then create a new slide to add text
                if(charInVerse + linesUsed > 9):
                    (slideLayout, slide, textBox) = self.setSlideInfo(1, title)
                    textBox = slide.shapes.placeholders[1]
                    textBox.text = lines
                    textFrame = textBox.text_frame
                    linesUsed = 1
                    continue
                #Add the verse to the frame
                verse = textFrame.add_paragraph()
                verse.text = lines
                linesUsed += charInVerse
        except OSError:#If the user enters an invalid book, tell them it is invalid
            print("Invalid Book")
        except ValueError:#If the user enters an invalid number, tell them it is invalid
            print("Must enter a valid number")
        except LookupError:#If the user enters an invalid verse, tell them it is invalid
            print("Invalid verses. No such passage exists")

    #Creates the header slide of the presentation (Helper Method)
    def header(self):
        #Gets the header information from the user
        print("Enter the header: ")
        userInput = self.getUserInput()
        #Assigns the values
        (slideLayout, slide, textBox) = self.setSlideInfo(2, userInput)

    #Defines the points of the message (Helper Method)
    def points(self):
        #Gets the title of the slides from the user
        print("Title of slide: ")
        userInput = self.getUserInput()

        #Creates the slide and textbox
        (slideLayout, slide, textBox) = self.setSlideInfo(1, userInput)
        #Assigns a placeholder for the frame
        textBox = slide.shapes.placeholders[1]
        textFrame = textBox.text_frame
        #Get all of the points from the user
        print("Enter the points for the slide. Type \'done\' to finish slide")
        textFrame.text = input("> ") + '\n'
        #Until they enter "done"
        while True:
            userInput = self.getUserInput()
            if userInput.lower().replace(" ", "") == 'done':
                break
            lines = textFrame.add_paragraph()
            lines.text = userInput + '\n'

    #Defines the lesson slides and asks the user for input each time until they quit
    def lessonSlides(self):
        while True:
            #B = Bible Passage, H = Header, P = Points, quit = create presentation
            print("\nChoices: \'B\' for Bible Passage, \'H\' for slide with a header, \'P\' for slide with points, \'quit\' to complete lesson")
            userInput = input("> ")
            userInput = userInput.upper().replace(" ", "")
            if(userInput == 'QUIT'):
                break
            elif(userInput == 'B'):#Create slide(s) with the passages
                self.biblePassage()
            elif(userInput == 'H'):#Create a header slide to define a section
                self.header()
            elif(userInput == 'P'):#Create a points slide to further discuss topic
                self.points()
            else:#Tell them it was an invalid choice
                print("Enter a valid choice")
        #Create and save the presentation
        self.presentation.save("your_powerpoint.pptx")