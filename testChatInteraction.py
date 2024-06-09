import socket
from time import sleep
from emoji import demojize

import win32api
import win32con
import win32gui

# Helper class for errors
class WindowNotFoundError(Exception):
    """Error to be raised if window name is not found"""
    pass

# Helper function grabbed at 2am, drank the win32 potion
def getWindowHwnd(window_name):
    """Will bring all windows with a specific name to the top. The last one to be brought up will be in focus.

    Todo:
        * I feel like I could do more with this. Might make it its own module

    Examples:
        >>> bring_window_to_top("Untitled - Notepad")
        # Will show all windows with the name "Untitled - Notepad"

    Args:
        window_name: (str) The name of the target window that you want to bring into focus

    Returns: None

    """

    def window_dict_handler(hwnd, top_windows):
        """Adapted from: https://www.blog.pythonlibrary.org/2014/10/20/pywin32-how-to-bring-a-window-to-front/

        """
        top_windows[hwnd] = win32gui.GetWindowText(hwnd)

    tw, expt = {}, True
    win32gui.EnumWindows(window_dict_handler, tw)
    for handle in tw:
        if tw[handle] == window_name:
            # print(tw[handle] + ": " + str(handle)) # This prints a handler, but we want the child window (for VBA specifically)
            return win32gui.ChildWindowFromPoint(handle, tuple([1, 1])) # As VBA has one child window, we will interact with this one (tweak this line for other programs)
    if expt:
        raise WindowNotFoundError(f"'{window_name}' does not appear to be a window.")

# Could theoretically produce hold functions if you were to apply this to another game
def makeCommand(hwndEdit, key):
    win32api.SendMessage(hwndEdit, win32con.WM_KEYDOWN, key, 0) # returns key sent
    sleep(0.1)
    win32api.SendMessage(hwndEdit, win32con.WM_KEYUP, key, 0) # returns key sent
    sleep(0.1)

# Key commands found here: https://learn.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
def interpretCommand(hwndEdit, comm):
    comm = comm.lower()
    if comm == 'v' or comm == 'down':
        makeCommand(hwndEdit, 0x53)
    elif comm == '^' or comm == 'up':
        makeCommand(hwndEdit, 0x57)
    elif comm == '<' or comm == 'left':
        makeCommand(hwndEdit, 0x41)
    elif comm == '>' or comm == 'right':
        makeCommand(hwndEdit, 0x44)
    elif comm == 'start':
        makeCommand(hwndEdit, 0x0d)
    elif comm == 'select':
        makeCommand(hwndEdit, 0x08)
    elif comm == 'screenshot':
        makeCommand(hwndEdit, 0x54)
    elif comm == 'a':
        makeCommand(hwndEdit, 0x5A)
    elif comm == 'b':
        makeCommand(hwndEdit, 0x58)
    elif comm == 'l':
        makeCommand(hwndEdit, 0x4C)
    elif comm == 'r':
        makeCommand(hwndEdit, 0x52)

# Use https://twitchapps.com/tmi/ to get your token for chat (oath pword)
conn_obj = {
    'server': 'irc.chat.twitch.tv',
    'port': 6667,
    'nickname': '', # Nickname is your username
    'token': 'oauth:7486u1aelnwng9x9a0l3l0ccas1r6b',
    'channel': '#' # channel is the chat room you'll connect to, e.g. #ninja
}

# Setup Twitch connection
sock = socket.socket()
sock.connect((conn_obj['server'], conn_obj['port']))

# Send the token, provide a username, and join the channel chat
sock.send(f"PASS { conn_obj['token'] }\n".encode('utf-8'))
sock.send(f"NICK { conn_obj['nickname'] }\n".encode('utf-8'))
sock.send(f"JOIN { conn_obj['channel'] }\n".encode('utf-8'))

# Setup VBA inputs
windowToFind = "emerald - VisualBoyAdvance-M 2.0.1" # Put the vba window name here
hwnd = getWindowHwnd(windowToFind) 
hwndEdit = win32gui.GetWindow(hwnd, win32con.GW_CHILD) # grabs the vba window

while True:
    try:
        resp = sock.recv(2048).decode('utf-8')

        if resp.startswith('PING'):
            sock.send("PONG\n".encode('utf-8'))

        elif len(resp) > 0:
            command = resp.split(conn_obj["channel"] + " :")[-1]
            command = demojize(command).strip()
            interpretCommand(hwndEdit,command)
    except:
        break

sock.close()