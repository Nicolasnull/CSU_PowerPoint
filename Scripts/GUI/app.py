from pptx import Presentation
from pptx.util import Pt
from tkinter import *
from CSUPowerPoint import *

class App:
    def main(interface):
        #Creates the variable to access the methods of the CSUPowerpoint module
        CSUPPT = CSUPowerPoint(interface)
        #Creates the introduction slide
        #CSUPPT.introduction()
        #Creates the title slide
        #CSUPPT.titleSlide()
        #Creates the officers slide
        #CSUPPT.officersSlide()
        #Creates the announcements slide
        #CSUPPT.announcementsSlide()
        #Creates the lesson header slides
        #CSUPPT.lessonHeaderSlide()
        #Creates the lesson slides
        #CSUPPT.lessonSlides()

    #If this is the file that is ran, then execute the main function
    if __name__ == "__main__":
        interface = Tk()
        main(interface)
        interface.mainloop()
