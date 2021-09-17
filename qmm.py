import pyautogui
import time
import win32api, win32con, win32clipboard
import random
import keyboard
import sys
# import tkinter as tk
from datetime import datetime
from pynput.keyboard import Key, Listener

# redirect print to file
old_stdout = sys.stdout
today = datetime.today().strftime('%Y-%m-%d')
log_file = open("QMM-"+today+".log","a", encoding="utf-8")
sys.stdout = log_file

# root = tk.Tk()
# keep the window from showing
# root.withdraw()

## Configuration
width = 1920
height = 1080
sleeprate = 0.01
scrollrate = 800
spamrate=12
status_exit=False
# Is print enabled?
print_enabled = True
# Output message
out_message=" "
# Set the toggle button
toggle_button = 'j'
exit_button = 'k'
# Set whether the script is enabled by default
enabled = False

## Utils
def quit():
	prnt(now() + " Exiting")
	sys.stdout = old_stdout
	log_file.close()
	sys.exit()

def now():
	return str(datetime.now().strftime("%H:%M:%S"))

def prnt(text):
	if (print_enabled):
		print(text)

def write(text):
	keyboard.write(text)
	time.sleep(sleeprate)

def press(key):
	keyboard.press(key)
	keyboard.release(key)
	time.sleep(sleeprate)

def get_copy():
	try:
		# return root.clipboard_get()
		# get clipboard data
		win32clipboard.OpenClipboard()
		data = win32clipboard.GetClipboardData()
		win32clipboard.CloseClipboard()
		time.sleep(sleeprate)
		return data
	except Exception as e:
		print(now() + " Exception occurred while copying data from clipboard: " + e)

def spam_rs(count=5):
	while count:
		write("say /rs")
		press("enter")
		count-=1
		time.sleep(sleeprate)

def set_cursor_pos(x,y):
	win32api.SetCursorPos((x, y))
	time.sleep(sleeprate*2)

def left_button_down(x,y,sleeprate=sleeprate):
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
	time.sleep(sleeprate*2)

def left_button_up(x,y,sleeprate=sleeprate):
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
	time.sleep(sleeprate*2)

def left_click():
	x,y = win32api.GetCursorPos()
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
	time.sleep(sleeprate*2)

def right_click():
	x,y = win32api.GetCursorPos()
	win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,x,y,0,0)
	win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,x,y,0,0)
	time.sleep(sleeprate*2)

def scroll(mag):
	x,y = win32api.GetCursorPos()
	win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, x, y, mag, 0)
	time.sleep(sleeprate)

# def move_mouse_relative(offX,offY):
# 	win32api.mouse_event(0x0001, offX, offY)

def copy_console_text_with_mouse(x,y):
	# open console
	press("`")

	# flood the chat (so that the chat appears in console)
	spam_rs(count=spamrate)

	# move mouse to start pos
	set_cursor_pos(x,y)
	scroll(int(0.75*scrollrate))

	# select text
	left_button_down(x,y)
	scroll(-scrollrate*2)
	set_cursor_pos(int(width*.75), int(height*.75))
	left_button_up(x,y)

	# move mouse to start pos
	set_cursor_pos(x,y)

	# copy
	right_click()
	time.sleep(sleeprate*10)
	left_click()
	time.sleep(sleeprate*10)

	# close console
	press("`")

def is_mouse_down():    # Returns true if the left mouse button is pressed
	lmb_state = win32api.GetKeyState(0x01)
	return lmb_state < 0

## Main functions
def print_startup():
	text = "Currently ENABLED" if enabled else "Currently DISABLED"
	prnt(now() + " Quick-Mafs-Machine Started! - " + text)

def say_console(text):
	global out_message
	if (text is not None and text != "None"):
		press("`")
		write("say " + text)
		press("enter")
		press("`")
		out_message = "result was \'" + text + "\'"
	else:
		out_message = "Exception: result was \'"+ text +"\'. See the logs for more info."
	prnt(now() + " output message: " + out_message)

def process_copy(c):
	prnt(now() + " Processing the copied Text (in reverse):")
	arr = c.split("\n")
	for item in arr[::-1]:
		prnt(item)
		# if (len(item)>4 and item[0]==":" and (item[-3:].lower()=="= ?" or item[-4:].lower()=="= ? ")):
		if (len(item)>4 and "=" in item and "?" in item):
			try:
				expr = item[1:-4].lower().replace(":","/").replace("x","*")
				return int(eval(expr))
			except:
				pass

def run():
	copy_console_text_with_mouse(int(0.065*width),int(0.17*height))
	c = get_copy()
	evaluated = process_copy(c)
	say_console(str(evaluated))

last_state = False
enabled = False
def onKeyPress(key):
	global last_state, enabled

	if ('char' not in dir(key)):
		return
	if (key.char==exit_button):
		quit()

	# print('You pressed %s\n' % (key, ))
	key_down = key.char == toggle_button
	if key_down != last_state:
			last_state = key_down
			if last_state:
				enabled = not enabled
				if enabled:
					prnt(now() + " Machine ENABLED")
					run()
					enabled = False
				prnt(now() + " Machine DISABLED")

def on_release(key):
	pass
    # print('{0} release'.format(key))
    # if key == Key.esc:
    #     return False

with Listener(on_press=onKeyPress, on_release=on_release) as listener:
    listener.join()