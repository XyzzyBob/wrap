#
# Wrap.py -- implement Windows mouse wrap-around on single or multiple monitors to allow wrapping at   
#   left, right or top, bottom borders.  Mouse will appear move to the opposite border and continue along it's 
#   merry old way.
#
# Version: 1.0.0
#
import sys
import pyWinhook as pyHook, mouse, ctypes, os, time, argparse, socket, select
import pythoncom
import win32.lib.win32con as win32con
import win32.win32gui as win32gui, win32.win32api as win32api

class Tracker:

    lSCREEN_WIDTH = lSCREEN_HEIGHT = 0		# local screen size, horizontal and vertical
    rSCREEN_WIDTH = rSCREEN_HEIGHT = 0		# remote screen size, horizontal and vertical
    SCREEN_WIDTH = SCREEN_HEIGHT = 0		# overall screen size, horizontal and vertical
    XMIN = XMAX = YMIN = YMAX = 0			# Active screen boundaries, dependent on which monitors we have
    lXMAX = lYMAX = lXMIN = lYMIN = 0		# Local screen boundaries, set based on local monitor stats
    lXPos = lYPos = rXPos = rYPos = 0       # Local and Remote mouse positions
    l_monArray = dict()
    
    def __init__(self):
        self.lXPos, self.lYPos = mouse.get_position()
        
    def local_move(self, x, y):
        mouse.move(x, y, absolute=True)
        self.XPos = x
        self.YPos = y

#
# Set up the argument parser stuff to handle arguments given on the command line
#
parser = argparse.ArgumentParser(description = "Cursor wrap for multiple screens")
parser.add_argument('-v', '--vertical',	required = False,	action="store_true", help="Monitors are arranged in a vertical stack rather than arranged horizontally")
parser.add_argument('-d', '--debug', required = False,  action="store_true",help="Debug mode -- prints statistics and information while running.")
args    = parser.parse_args()

vert = True if args.vertical else False     # Are we running in vertical or horizontal mode? Default is horizontal
debug = True if args.debug else False       # If set to true, additional stats and info are provided when running

def onBorder(xpos, ypos):
    if (progParam.XMIN < xpos < progParam.XMAX) and (progParam.YMIN < ypos < progParam.YMAX):
        return 'no'
    else:    
        return 'left' if xpos <= (progParam.XMIN) else \
            'top' if (ypos <= progParam.YMIN) else \
            'right' if (xpos >= progParam.XMAX) else 'bottom'
# 
# Routine to handle screen-wrap when we hit an edge.  Set the position to 1 more or 1 less than the opposing edge
#
def setWrapPos(xpos, ypos, border):
	if border == 'left' or border == 'right':
		xpos = progParam.XMAX-1 if border == 'left' else progParam.XMIN+1
	elif border == 'top' or border == 'bottom':
		ypos = progParam.YMAX - 1 if border == 'top' else progParam.YMIN + 1
	else:
		print('ERROR: Illegal border passed to setWrapPos()')

	mouse.move(xpos, ypos, absolute=True)
	return True

def onclick(event):
	print(event.event_type + ' for ' + event.button + ' button at ' + str(mouse.get_position()))
	return True

def mouseEvent(event):
	if isinstance(event, mouse._mouse_event.MoveEvent):
		trackMouse(event)
	return True

def trackMouse(event):
    xpos = event.x
    ypos = event.y
    border = 'no'

    for key, rec in progParam.l_monArray.items():
        if vert: 
            if rec[2] < ypos < rec[4]:
                progParam.XMIN = rec[1]
                progParam.XMAX = rec[3]
                break
        else:
            if rec[1] < xpos < rec[3]:
                progParam.YMIN = rec[2]
                progParam.YMAX = rec[4]
    if debug: print("Mouse at (%d, %d)" % (xpos, ypos))
    border = onBorder(xpos, ypos)
    if border != 'no':
        setWrapPos(xpos, ypos, border)
#    if debug: print("Boundaries:  left - " + str(progParam.XMIN) + ", right - " + str(progParam.XMAX) + ", top - " + str(progParam.YMIN) + ", bottom - " + str(progParam.YMAX))
    return True

def onWinCombo(event):
	if event.Key == 'End':
		print('CTRL-End was detected.  Exiting.')
		hm.UnhookMouse()
		hm.UnhookKeyboard()
		os._exit(0)
	else:
		hm.KeyDown = onKeyboardEvent
	return True

def cancelCombo(event):
	hm.KeyDown = onKeyboardEvent
	hm.KeyUp = None
	return True

def onKeyboardEvent(event):
    if debug: print('Keyboard action detected, ' + str(event.Key) + ' was pressed.')
    if str(event.Key) == 'Lcontrol' or str(event.Key) == 'Rcontrol':
        hm.KeyDown = onWinCombo
        hm.KeyUp = cancelCombo
        return True
    else:
        return True

progParam = Tracker()

mons = win32api.EnumDisplayMonitors()
i = 0
while i < win32api.GetSystemMetrics(win32con.SM_CMONITORS):
    minfo = win32api.GetMonitorInfo(mons[i][0])
    progParam.l_monArray[i] = ([minfo['Monitor'][2] - minfo['Monitor'][0],
        minfo['Monitor'][3] - minfo['Monitor'][1]], minfo['Monitor'][0], minfo['Monitor'][1],
        minfo['Monitor'][2], minfo['Monitor'][3])
    i += 1 
    if debug: print(minfo)
for key, rec in progParam.l_monArray.items():
    if vert:
        progParam.lSCREEN_HEIGHT += rec[0][1]
        progParam.lSCREEN_WIDTH = (rec[0][0]) if progParam.lSCREEN_WIDTH < rec[0][0] else progParam.lSCREEN_WIDTH
        progParam.lYMAX = (rec[4]) if progParam.lYMAX < rec[4] else progParam.lYMAX
        progParam.lYMIN = (rec[2]) if progParam.lYMIN > rec[2] else progParam.lYMIN
        progParam.lXMAX = (rec[3]) if progParam.lXMAX < rec[3] else progParam.lXMAX
        progParam.lXMIN = (rec[1]) if progParam.lXMIN > rec[1] else progParam.lXMIN
    else:
        progParam.lSCREEN_WIDTH += rec[0][0]
        progParam.lSCREEN_HEIGHT = (rec[0][1]) if progParam.lSCREEN_HEIGHT < rec[0][1] else progParam.lSCREEN_HEIGHT
        progParam.lXMAX = (rec[3]) if progParam.lXMAX < rec[3] else progParam.lXMAX
        progParam.lXMIN = (rec[1]) if progParam.lXMIN > rec[1] else progParam.lXMIN
        progParam.lYMAX = (rec[4]) if progParam.lYMAX < rec[4] else progParam.lYMAX
        progParam.lYMIN = (rec[3]) if progParam.lYMIN > rec[3] else progParam.lYMIN
    if debug: print("Monitor #" + str(key) + ": " + str(rec[0][0]) + "x" + str(rec[0][1]))
if debug: print(progParam.l_monArray)

progParam.SCREEN_HEIGHT = progParam.lSCREEN_HEIGHT
progParam.SCREEN_WIDTH = progParam.lSCREEN_WIDTH
progParam.XMAX = progParam.lXMAX
progParam.XMIN = progParam.lXMIN
progParam.YMAX = progParam.lYMAX
progParam.YMIN = progParam.lYMIN

if debug: print("XMAX: " + str(progParam.lXMAX) + " XMIN: " + str(progParam.lXMIN) + " YMAX: " + str(progParam.lYMAX) + " YMIN: " + str(progParam.lYMIN))
if debug: print("Screen Height: " + str(progParam.SCREEN_HEIGHT) + " Screen Width: " + str(progParam.SCREEN_WIDTH))

mouse_hm = mouse.hook(mouseEvent)
print("\nWrap() starting -- Use ctrl+End to exit.\n")
hm = pyHook.HookManager()
hm.KeyDown = onKeyboardEvent
hm.HookKeyboard()
pythoncom.PumpMessages()
