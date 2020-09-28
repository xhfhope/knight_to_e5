import pyautogui
import speech_recognition as sr
import time
r = sr.Recognizer()
m = sr.Microphone()
from pynput import mouse
import win32gui
import win32con
from os import system, name 
cHwnd = 0

def clearConsole():
	_ = system('cls')

def winEnumHandler( hwnd, ctx ):
	global cHwnd
	if win32gui.IsWindowVisible( hwnd ):
		if 'cmd.exe' in win32gui.GetWindowText(hwnd):			
			if 'chessmic' in win32gui.GetWindowText(hwnd):				
				cHwnd = hwnd

win32gui.EnumWindows( winEnumHandler, None )
rect = win32gui.GetWindowRect(cHwnd)
win32gui.SetWindowPos(cHwnd, win32con.HWND_TOPMOST, 0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

#======================================================
#				   RECEIVE BOARD BOUNDS
#======================================================
print('\nYour next two mouse clicks need to be the center of the bottom left square and the center of the top right square, in that order\n')
startpointX = 0
startpointY = 0
endpointX = 0
endpointY = 0

def on_click(x, y, button, pressed):
	global startpointX, startpointY, endpointX, endpointY
	if button == mouse.Button.left and pressed:		
		if startpointX==0:
			startpointX, startpointY = pyautogui.position()
		else:
			endpointX, endpointY = pyautogui.position()		
		
listener = mouse.Listener(on_click=on_click)
listener.start()

while endpointX ==0:
	pass
listener.stop()



#======================================================
#				   DEFINE ALL COORDINATES
#======================================================
#speechDict = {0:'alpha ', 1:'beta ', 2:'charlie ', 3:'delta ', 4:'epsilon ', 5:'foxtrot ', 6:'gamma ', 7:'hotel '}
speechDict = {0:'a', 1:'b', 2:'c', 3:'d', 4:'epsilon', 5:'f', 6:'g', 7:'h'}
width = (endpointX - startpointX) / 7
height = (endpointY - startpointY) / 7
Coords = {}
i = 0
j =0
while i<8:
	j=0
	while j<8:
		x,y = (startpointX + width*j, startpointY + height*i)
		s = speechDict[j] + str(i+1)
		Coords.update( {s:(x,y)})
		j=j+1
	i=i+1
#print(Coords)
pyautogui.click(Coords['d4'])

firstPass = True
#======================================================
# 						MAIN LOOP
#======================================================
try:
    #clearConsole()
    print("A moment of silence, please...")
    while True:
        found = []        
        try:            
            with m as source: r.adjust_for_ambient_noise(source)       
            if firstPass:
        	    print('Ready')
        	    firstPass = False
            with m as source: audio = r.listen(source)
            value = r.recognize_google(audio)
            value = value.lower()
            if value == 'new opponent':
                pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('new_opponent.png')))
            elif value == 'rematch':
                pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('rematch.png')))
            elif value == 'home':
            	pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('home.png')))
            elif value == '10 minute' or value=='ten minute':
            	pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('10min.png')))
            elif value == '5 minute' or value=='five minute':
            	pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('5min.png')))
            elif value == '15 minute' or value=='fifteen minute':
            	pyautogui.click(pyautogui.center(pyautogui.locateOnScreen('15min.png')))
            value = value.replace('-','')
            value = value.replace('for','4')
            value = value.replace('four','4')
            value = value.replace('to','2')
            value = value.replace('too','2')
            #clearConsole()
            print('\nReceived "{}"!'.format(value))            
            for key in Coords:
                if key in value:
                   found.append(key)
            if len(found) == 2:
                if found[0] in value[0:len(found[0])]:
                	i = 0
                	j = 1
                else:
                	i = 1
                	j = 0
                x,y=Coords[found[i]]
                pyautogui.moveTo(x,y,0.5,pyautogui.easeOutQuad)
                pyautogui.mouseDown();
                x,y=Coords[found[j]]
                pyautogui.moveTo(x,y,0.5,pyautogui.easeOutQuad) 
                pyautogui.mouseUp();
                pyautogui.moveTo(0,0,0.5,pyautogui.easeOutQuad) 
                #pyautogui.click(Coords[found[i]])
                #time.sleep(0.5)
                #pyautogui.click(Coords[found[j]])
        except sr.UnknownValueError:
            print(".", end='')
        except sr.RequestError as e:
            print("Couldn't request results from Google Speech Recognition service; {0}".format(e))
except KeyboardInterrupt:
	print('Exit.')
	win32gui.SetWindowPos(cHwnd, win32con.HWND_NOTOPMOST, 0,0,0,0,win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)