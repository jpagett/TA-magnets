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

infoTable = "info.xlsx"

FullNameCol = 0
# infoTable format: first (index 0) column as FullNames (First Last)

CourseNumCol = 1
# infoTable format: second (index 1) column as Course Numbers (PHYS 103)

template = "nametagTemplate.png"

log = open(os.path.join(script_dir,outputFolder,"errors.txt"),"w")

def makeTag(FullName, CourseNum):

    multCourse = False
    # workaround -- TAs are for one course until found otherwise

    nameX, nameY = 180, 627
    # pixel location of the name's upper-left corner

    courseX, courseY = 728, 255
    # pixel location of the course number (UL corner)

    imageX, imageY = 100, 100
    # pixel location of the picture
    # auto shifts for certain picture dimensions to attempt to center (line 85-92)
    # very dumb, should rework to center on desired area instead of fixed shifts

    pic_width, pic_height = 500, 450
    # sets maximum dimensions (scales other to maintain aspect ratio)

    FullName = str(FullName)
    First, Last = FullName.split(" ")

    font = ImageFont.truetype(os.path.join(fontFolder,Font),100)

    imageFile = os.path.join(script_dir,template)
    outputPath = os.path.join(outputFolder, f'{First}{Last}_tag.png')
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

    pictureName = '{}{}.jpg'.format(First,Last)
    picturePath = os.path.join(script_dir,pictureFolder,pictureName)
    picture = Image.open(picturePath)
    # puts together path for and opens picture

    width, height = picture.size

    if width > height:
        pic_height = int(pic_width * height / width)
        imageY += 25
        if pic_height < 300:
            imageY += 50
    else:
        pic_width = int(pic_height * width / height)
        imageX += 50

    picture = picture.resize((pic_width, pic_height))

    # this clusterfuck decides the size of the pictures, depends on whether
    # portrait or landscape picture

    if multCourse:
        imageX = imageX - 40
    # shifts image over if TA covers 2 courses

    im1.paste(picture,(imageX,imageY))
    # pastes picture on template

    im1.save(outputImage)
    # saves tag

book = xlrd.open_workbook(os.path.join(script_dir,infoTable))
sheet = book.sheet_by_index(0)
# loads the excel doc of TA names and courses
# should be formatted as follows:

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
