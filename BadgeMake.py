import PIL
from PIL import ImageFont
from PIL import Image
from PIL import ImageDraw

import os.path
import xlrd
import xlsxwriter
from xlrd import open_workbook

script_dir = os.path.dirname(__file__)

pictureFolder = 'Pictures'
outputFolder = 'Output'
# these two folders should be created within the folder containing this script

fontFolder = "C:/Windows/fonts"
Font = "calibri.ttf"

infoPath = "info.xlsx"

log = open(os.path.join(script_dir,outputFolder,"errors.txt"),"w")

def makeTag(FullName, CourseNum):

    multCourse = False
    # workaround -- TAs are for one course until found otherwise

    nameX = 180
    nameY = 627
    # pixel location of the name's upper-left corner

    courseX = 728
    courseY = 255
    # pixel location of the course number (UL corner)

    FullName = str(FullName)
    First, Last = FullName.split(" ")

    font = ImageFont.truetype(os.path.join(fontFolder,Font),100)

    imageFile = os.path.join(script_dir,"nametagTemplate.png")
    outputPath = os.path.join(outputFolder,First + Last + "_tag.png")
    outputImage = os.path.join(script_dir,outputPath)
    im1 = Image.open(imageFile)

    if len(FullName) > 22:
        nameX = nameX - 120
    #shifts name over if too long. Only works for len(FullName =< 22)

    Name = ImageDraw.Draw(im1)
    Name.text((nameX,nameY),FullName,(0,0,0),font=font)
    #Name = ImageDraw.Draw(im1)
    # places name on image

    im1.save(outputImage)
    im1 = Image.open(outputImage)
    # saves working instance with text added

    if len(CourseNum.split(" ")) > 2:
        courseX = courseX - 150
        multCourse = True
    # if TA serves 2 courses, shifts course placement over to make room

    Course = ImageDraw.Draw(im1)
    Course.text((courseX,courseY),CourseNum,(0,0,0),font=font)
    im1.save(outputImage)
    im1 = Image.open(outputImage)
    # adds course number text, saves as working instance

    pictureName = First + Last + '.jpg'
    picturePath = os.path.join(script_dir,pictureFolder,pictureName)
    picture = Image.open(picturePath)
    # puts together path for and opens picture

    width, height = picture.size

    if width > height:
        new_width  = 500
        new_height = int(new_width * height / width)
        imageX = 100
        imageY = 125
        if new_height < 300:
            imageY = 175
    else:
        new_height = 400
        new_width = int(new_height * width / height)
        imageX = 150
        imageY = 100

    picture = picture.resize((new_width, new_height))

    # this clusterfuck decides the size of the pictures, depends on whether
    # portrait or landscape picture

    if multCourse:
        imageX = imageX - 40
    # shifts image over if TA covers 2 courses

    im1.paste(picture,(imageX,imageY))
    # pastes picture on template

    im1.save(outputImage)
    # saves tag

book = xlrd.open_workbook(os.path.join(script_dir,infoPath))
sheet = book.sheet_by_index(0)
# loads the excel doc of TA names and courses
# should be formatted as follows:

FullNameCol = 0
# first column as FullNames (First Last)

CourseNumCol = 1
# second column as Course Numbers (PHYS 103)

for k in range(1,sheet.nrows):

        FullName = sheet.row_values(k)[FullNameCol]
        Course = sheet.row_values(k)[CourseNumCol]

        try:
            makeTag(FullName,Course)
        # makes tags

        except:
            log.write(FullName+"\n")
        # writes TA names which cause errors to errors.txt

log.close()
