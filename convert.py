import cv2
import glob
import os
from pptx import Presentation
from pptx.util import Inches
import sys

# Ensuring valid argument
if len(sys.argv) != 2:
    sys.exit("Invalid argument (video path should be the only argument)")

videofile = sys.argv[1]

if videofile[-4:] != ".mp4":
    sys.exit("Invalid format (format should be .mp4)")

if not os.path.isfile(videofile):
    sys.exit("File does not exist")

global frames
frames = os.path.dirname(videofile) + "\\frames\\"

# Function to set up path
def delete_frames(createdir):
    try:
        if os.path.exists(frames):
            for filename in glob.glob(frames + "*"):
                os.remove(filename)  # Deletes each file in \frames\
            os.rmdir(frames)
        if createdir:
            os.mkdir(frames)
    except OSError as error:
        sys.exit(error)


delete_frames(True)

# Initializing
vid = cv2.VideoCapture(videofile)
framerate = vid.get(cv2.CAP_PROP_FPS)

# Reading video
framenumber = 0
while True:
    framesleft, frame = vid.read()
    if not framesleft:
        break
    else:
        cv2.imwrite(f"{frames}{framenumber}.jpg", frame)
        framenumber += 1
vid.release()

# Creating pptx
prs = Presentation()
prs.slide_width = Inches(16)
prs.slide_height = Inches(9)
totalframes = len(glob.glob(frames + "*"))
freq = framerate / 30
framenumber = 0
while int(framenumber) <= totalframes - 1:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(f"{frames}{int(framenumber)}.jpg", 0, 0, width=Inches(16), height=Inches(9))
    framenumber += freq

# Clean-up
prs.save(f"{videofile[:-3]}pptx")
delete_frames(False)
