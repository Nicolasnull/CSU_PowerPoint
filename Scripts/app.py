from pptx import Presentation
from pptx.util import Pt
from CSUPowerPoint import *

class App:
    def main():
        CSUPPT = CSUPowerPoint()
        CSUPPT.introduction()
        CSUPPT.titleSlide()
        CSUPPT.officersSlide()
        CSUPPT.announcementsSlide()
        CSUPPT.lessonHeaderSlide()
        CSUPPT.lessonSlides()

    if __name__ == "__main__":
        main()
