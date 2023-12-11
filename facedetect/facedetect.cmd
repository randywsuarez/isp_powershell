@echo off
echo ===Command line for facedetect.exe===

facedetect --cascade="haarcascade_frontalface_alt.xml" --CamIndex=0 --duration=3 --jpg --faceCount=20 --lefttop_x=240 --lefttop_y=90 --rightbottom_x=400 --rightbottom_y=270 > result.ini