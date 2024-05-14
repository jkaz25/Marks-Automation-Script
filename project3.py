# Assignment: project3.py
# Programmer: Joey Kaz
# Date: May 6, 2024
# This Python script is used to emulate a production workflow
#The program uses argparse for user input and performs the following primary operations:
# 1) imports text file data from a Xytech workorder and populates a database collection 
# 2) Imports frame information from a baselight output and populates a database collection with the frame ranges and file locations
# 3) Uses third party tool ffprobe to extract video duration and fps 
# 4) Read data from populated databases and swap correct file ranges
# 5) Calculate timecode from marks on baselight output
# 6) Read marks from baselight data to extract images from video theat fall within the range of duration
# 7) Extract the timecode range for each frame range
# 8) Upload frame images to FrameIO project folder via FrameIOClient API


#import libraries 
import argparse as ap
import pandas as pd
import pymongo as pm
import subprocess as sp
from frameioclient import FrameioClient
import xlsxwriter as x

# global dictionaries to keep track of imported data

baselightData = {
    "files": [],
    "shots" :[]
}

XytechData = {}

#argparse argument declarations and call
parser = ap.ArgumentParser()
parser.add_argument('-p', '--process' , nargs=1, dest='process')
parser.add_argument('-b' , '--baselight', action='store_true', dest='baselight')
parser.add_argument('-xy' , '--xytech' , action='store_true', dest='xytech')
args = parser.parse_args()

#establish connection to mongodb localhost and create collections for data
myclient = pm.MongoClient('mongodb://localhost:27017/')
mydb = myclient['project3']
#if 'baselight' in mydb.list_collection_names():
    #mydb.drop_collection('baselight')
#if 'xytech' in mydb.list_collection_names():
    #mydb.drop_collection('xytech')
blCollection = mydb['baselight']
xytechCollection = mydb['xytech']

#establishes connection to frameio host and project
client = FrameioClient("fio-u-OJ4wH42Spv5htxLkYAkAyxf9C8T_Mq45BAAzUdLI8QmL49PlTmHDZdnepuenWJkc")
the_crucible = client.projects.get("e75fe748-5ec9-48a7-a366-20f07d9546fb")

#set up spreadsheet workbook and main worksheet
wb = x.Workbook('project3.xlsx')
ws = wb.add_worksheet()
#set column width for images
ws.set_column(3,3,27)
ws.set_column(2,2,21)
ws.set_column(1,1,14)
ws.set_column(0,0,50)
ws.set_row(4,96)

#parses text file for shotlist and returns their file location + list
def parseShots(line):
    line = list(line)
    key = line[0]
    shots = []
    for i in range(1, len(line)):
        if str(line[i]).isnumeric():
            shots.append(line[i])
    return key, shots

#method that specifically removes the split file path
def remove_filePath(line):
    line = list(line)
    return line[1].split(" ")

#method that processes the baselight export file
def processBL():
    # open reading connection to file
    with open("Baselight_export.txt", "r") as bl_file:
        baselight = bl_file.readline()
        #parse file until no more data
        while baselight != "":
            #check in case empty lines in file to avoid processing errors
            if baselight.rstrip("\n") == "":
                baselight = bl_file.readline()
                continue
            #split each line, remove file location, create shot list, and then store it in dictionary
            bl = baselight.rstrip("\n").split("Dune2")
            bl = remove_filePath(bl)
            #get the shotlist for each specified filepath
            key, shots = parseShots(bl)
            blCollection.insert_one( {
                "file" : key,
                "shotList" : shots
            })
            baselight = bl_file.readline()
            # loop to read data until eof
    #close file connection
    bl_file.close()

#method that processes the xytech work order file
def process_Xytech():
    #open file for reading
    with open("Xytech.txt", 'r') as xytech_file:
        xy = xytech_file.readline()
        # loop until end of file
        while xy != "":
            #because there are empty lines before eof, check to avoid processing errors
            if xy.rstrip("\n") == "":
                xy = xytech_file.readline()
                continue
            #checking for individual fields to store for later export
            if "Producer" in xy or 'Operator' in xy or 'Job' in xy:
                xy = xy.rstrip("\n").split(":")
                xytechCollection.insert_one({xy[0].lstrip() : xy[1]})
            if "Dune2" in xy:
                xy = xy.rstrip("\n").split("Dune2")
                xytechCollection.insert_one({xy[0] : xy[1]})
            if "Notes" in xy:
                xy = xy.rstrip("\n")
                xytechCollection.insert_one({"Notes" : xytech_file.readline().rstrip("\n")})
            xy = xytech_file.readline()
        #close file connection
        xytech_file.close()

#checks if file location from xytech workorder already exists in dictionary
def check_Xytech(line):
    if line in XytechData.keys():
        return True
    return False

#method that queries the dataset from the baselight collection
def queryBaselight():
    df = pd.DataFrame(blCollection.find({},{"_id" : 0}))
    #append file names to baselight data
    for file in df['file']:
        baselightData['files'].append(file)
    #append shot lists to baselight data
    for shotList in df['shotList']:
        #convert shots to integer type
        baselightData['shots'].append([int(x) for x in shotList])

#method that queries the dataset from the xytech collection
def queryXytech():
    df2 = list(xytechCollection.find({}, {"_id":0}))
    checks = ['Producer', 'Operator', 'Job','Notes']
    for i in df2:
        #since each result from the database is returned as a dictionary, go through each one and turn the key/value into a list for parsing
        key = list(i.keys())
        value = list(i.values())
        if key[0] in checks:
            XytechData[key[0]] = value[0]
        else:
            if check_Xytech(key[0]) == True:
                    #if location exists, append new scene
                    XytechData[key[0]].append(value[0])
            else:
                # if doesn't exist, create new scene list for location
                XytechData[key[0]] = value

#retrieves the file lcocation from the xytech dictionary and returns the path + Dune
def get_fileLocation(key):
    for newLocation in XytechData.keys():
        for file in XytechData[newLocation]:
            if key == file:
                return newLocation + "Dune2"
           
#method that calculates timecode
def getTimeCode(fps: int, frame: int):
    tSeconds, frames = frame//fps, frame%fps
    tMinutes, seconds = tSeconds//60, tSeconds%60
    tHours, minutes = tMinutes//60, tMinutes%60
    hours = tHours%60
    segments = [hours, minutes, seconds, frames]
    #pads each timecode with a zero segment if less than 10
    paddedSegments = [''.join(('0', str(segment))) if segment < 10 else str(segment) for segment in segments]
    #returns formatted timecode
    return f"{paddedSegments[0]}:{paddedSegments[1]}:{paddedSegments[2]}:{paddedSegments[3]}"

#method that takes screenshot of specified timestamp from timecode and returns image filepath for upload
def getImage(timestamp, count, videofile): 
    parts = timestamp.split(":")
    #converts the remaining frames section of TC to fractions of a second
    #parts[3] = the frames section of TC
    fraction1 = str(float(parts[3])/60)[1:] #[1:] removes the 0 before the decimal point for formatting
    fraction2 = str(float(parts[3])/60 + 0.01)[1:]
    newTC = f'{parts[0]}:{parts[1]}:{parts[2]}{fraction1}'
    newTC2 = f'{parts[0]}:{parts[1]}:{parts[2]}{fraction2}'
    #take screen shot with specified aspect ratio and timecode/timestamp
    sp.call(f'ffmpeg -i {videofile} -aspect 96:74 -ss: {newTC} -to {newTC2} -r 1 -f image2 image{count}.png')
    return f"image{count}.png"

#method that checks to see if a frame's timecode is in the duration for the video (excludes excess frames)
def isInRange(timecode,duration):
    #splits TC and video duration into parts for comparison
    timecode = timecode.split(":")
    duration = duration.split(":")
    #timecode within duration for hours
    if timecode[0] > duration[0]:
        return False
    #timecode within duration for minutes
    if timecode[1] > duration[1]:
        return False
    #timecode within duration for seconds
    if timecode[2] > duration[2]:
        return False
    return True

#method that uploads specified image to frameio project folder via API call
def upload(image):
    asset = client.assets.upload (
    destination_id=the_crucible['root_asset']['id'],
    filepath=f"{image}"
)
#method that writes xytech workorder headers to spreadsheet
def writeHeaders():
    ws.write("A1", "Producer")
    ws.write("B1", "Operator")
    ws.write("C1", "Job")
    ws.write("D1", "Notes")
    ws.write("A2",XytechData["Producer"])
    ws.write("B2", XytechData["Operator"])
    ws.write("C2", XytechData["Job"])
    ws.write("D2", XytechData["Notes"])
    ws.write("A4", "Show Location")
    ws.write("B4", "Frames to Fix")
    ws.write("C4", "Timecode Ranges")
    ws.write("D4", "Screenshots")

#method that writes a range of frames to spreadsheet
def writeRange(row: int, file: str, key: str, slIndex1: int, slIndex2: int, tcIndex1: int, tcIndex2: int, shotList: list, fps: int, imagepath: str) -> None:
    ws.write(f"A{row}", f"{file}{key}")
    ws.write(f"B{row}", f"{shotList[slIndex1]}-{shotList[slIndex2]}")
    ws.write(f"C{row}",f"{getTimeCode(fps, shotList[tcIndex1])}-{getTimeCode(fps,shotList[tcIndex2])}")
    ws.set_row(row, 96)
    #set row height for imaage
    ws.insert_image(f"D{row}", imagepath,{'x_scale': 0.1, 'y_scale':0.1})

#method that writes a single frame to spreadsheet for shots to fix
def writeSingle(row: int, file: str, key: str, slIndex: int, tcIndex: int, tcIndex2: int, shotList: list, fps: int, imagepath: str) -> None: 
    ws.write(f"A{row}", f"{file}{key}")
    ws.write(f"B{row}", f"{shotList[slIndex]}")
    ws.write(f"C{row}",f"{getTimeCode(fps, shotList[tcIndex])}-{getTimeCode(fps,shotList[tcIndex2])}")
    #set row height for image
    ws.set_row(row,96)
    ws.insert_image(f"D{row}", imagepath, {'x_scale': 0.1, 'y_scale':0.1}) 

#method that finalizes the data export
def export(duration: int, fps: int, videofile: str):
    #open file for writing data
    row = 5 # starting value for export data
    writeHeaders()
    #writes meta data about producer, operator, job, and notes
    shotListIndex = 0
    imagecount = 0
    #loop through all file locations in baselight and grab new file location from Xytech
    for key in baselightData["files"]:
        shotList = baselightData["shots"][shotListIndex]
        #see if shotList is in duration by checking last frame
        timecode = getTimeCode(fps, int(shotList[-1]))
        #break out of loop if out of range
        if isInRange(timecode,duration) == False:
            break
        newFile = str(get_fileLocation(key))
        #loop through the shotlists for each files location
        counter = 1
        first = 0
        for i in range(1, len(shotList)):
            if shotList[i] > shotList[i-1] + 1:
                #determine last and middle frames to extract timecode and write to output
                last = i-1
                mid = (first + last) // 2
                imagecount = imagecount + 1
                timecode = getTimeCode(fps, int(shotList[mid]))
                imagepath = getImage(timecode,imagecount, videofile)
                if counter == 1:
                    writeSingle(row, newFile, key, first,first,first,shotList,fps,imagepath)
                else:
                    writeRange(row,newFile, key, first, last, first, last, shotList, fps,imagepath)
                    counter = 1
                first = i
                row = row + 1
                upload(imagepath)
            else:
                counter += 1
        #reach end of shot list
        mid = (first + i) // 2
        timecode = getTimeCode(fps, int(shotList[mid]))
        imagecount = imagecount + 1
        imagepath = getImage(timecode,imagecount,videofile)
        if counter > 1:
            writeRange(row, newFile, key, first, -1, first, i, shotList, fps,imagepath)
        else:
            writeSingle(row, newFile, key, i, first, i, shotList, fps,imagepath)
        upload(imagepath)
        row = row + 1
        shotListIndex+=1

#method that retrieves the duration and fps from the video input file
def getDurationAndFPS(videofile: str):
    duration = 0
    fps = 0
    x = sp.getoutput(f"ffprobe {videofile}")
    #store ffprobe call output as txt file
    with open("ffprobeOutput.txt", 'w') as op:
        op.write(x)
        op.close()

    #open text file output and parse for duration and fps
    with open('ffprobeOutput.txt', 'r') as op:
        line = op.readline()
        line = line.lstrip()
        while line.rstrip("\n") != "":
            if 'Duration' in line:
                line = line.split(',')
                line = line[0].lstrip()
                line = line.split(" ")[1].lstrip()
                duration = line
            if "Stream" in line:
                line = line.split(",")
                for section in line:
                    if 'fps' in section:
                        fps = section.lstrip().split(" ")[0]
            line = op.readline()
        op.close()
    return duration, int(fps)

#main method 
def process():
    videofile = args.process[0]
    queryBaselight()
    queryXytech()
    duration, fps = getDurationAndFPS(videofile=videofile)
    export(duration, fps, videofile)

#argparse method calls
if args.baselight:
    processBL()

if args.xytech:
    process_Xytech()

if args.process:
    process()

#close connection to workbook if program executes successfully
wb.close()