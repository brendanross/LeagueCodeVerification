import Image
import ImageGrab
import cv2
from cv2 import cv
from win32com.client import Dispatch
import win32api, win32con
import time
import os

output = os.getcwd() + r'\OUTPUT' 
if not os.path.exists(output): 
	os.makedirs(output)

scr = 'screen.png'
method = cv.CV_TM_SQDIFF_NORMED
SendKeys = Dispatch("WScript.Shell").SendKeys
with open('codes.txt') as f:
		codes = f.readlines()

def click(x,y):
	x+=4
	y+=4
	win32api.SetCursorPos((x,y))
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

def grabScreen():
	img = ImageGrab.grab()
	img.save("screen.png")

def compareImages(small):
	small = cv2.imread(small)
	large = cv2.imread(scr)
	result = cv2.matchTemplate(small, large, method)
	mn,_,mnLoc,_ = cv2.minMaxLoc(result)
	MPx,MPy = mnLoc
	return MPx, MPy

def goToStore():
	grabScreen()
	time.sleep(2)
	MPx,MPy = compareImages('img\\storeButton.png')
	click(MPx, MPy)
	
def scrollDown(scr):
	MPx,MPy = compareImages('img\\downButton.png')
	x=0
	while x<4:
		click(MPx, MPy)
		x+=1
		
def goToCodes():
	grabScreen()
	time.sleep(2)
	scrollDown(scr)
	time.sleep(1)
	grabScreen()
	MPx,MPy = compareImages('img\\codeButton.png')
	click(MPx, MPy)

def inputCode(code):
	SendKeys(code.replace('\n',''))
	time.sleep(1)
	grabScreen()
	MPx,MPy = compareImages('img\\submitButton.png')
	click(MPx,MPy)
	
def goToTextField():
	time.sleep(2)
	grabScreen()
	MPx,MPy = compareImages('img\\inputField.png')
	click(MPx, MPy)
	SendKeys('^a')
	SendKeys('{DELETE}')

def clickX():
	MPx,MPy = compareImages('img\\xButton.png')
	click(MPx, MPy)
	time.sleep(2)

#0=valid, 1=invalid, 2=used
def checkValid():
	grabScreen()
	states = cv2.imread('img\\validText.png'), cv2.imread('img\\invalidText.png'), cv2.imread('img\\usedText.png')
	large = cv2.imread(scr)
	x = 0
	returnCode = []
	while x < 3:
		result = cv2.matchTemplate(states[x], large, method)
		mn,_,mnLoc,_ = cv2.minMaxLoc(result)
		returnCode.append(mn)
		x+=1
	return returnCode.index(min(returnCode))
		
		


goToStore()
time.sleep(2)
goToCodes()
time.sleep(2)

for code in codes:
	goToTextField()
	time.sleep(1)
	inputCode(code)
	time.sleep(2)
	value = checkValid()
	if value == 0:
		#valid
		with open('OUTPUT\\VALID.txt','a') as file:
			file.write(str(code))
			file.close()
	elif value == 1 or value == 2:
		#invalid
		with open('OUTPUT\\INVALID.txt','a') as file:
			file.write(str(code))
			file.close()

	clickX()
	time.sleep(2)
print "Check OUTPUT folder for lists of valid/invalid codes"
os.system('pause')