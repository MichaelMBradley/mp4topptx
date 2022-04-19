# mp4topptx

A script that takes videos in an mp4 format and converts them into a PowerPoint presentation.
Requires the `opencv-python` and `python-pptx` packages.

To use, run in the command prompt (or in a batch file):

```lang-none
python3 convert.py [path to target mp4]
```

The program will briefly make a directory of the frames of the mp4, but will then delete it automatically *(Update 22/04/19: accidentally broke this functionality)*. It will generate a powerpoint file with the same name and path as the mp4, but with a .pptx extension.

I made some changes recently to make it a bit more pythonic and proper but it's still pretty rough.
