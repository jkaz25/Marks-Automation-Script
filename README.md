# Marks-Automation-Script
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
