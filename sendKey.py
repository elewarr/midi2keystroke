__author__ = 'Nick Mullen AKA CMDR Eld Ensar'

## This is mostly the work of others...
## the bulk of it being lifted from user Hodka at stackoverflow, who had the best version I saw of sending keyboard ScanCodes rather than VK codes
## his original work is here:   http://stackoverflow.com/a/23468236


import ctypes
import json
import time
import datetime
import winsound
import configparser

SendInput = ctypes.windll.user32.SendInput

# C struct redefinitions
PUL = ctypes.POINTER(ctypes.c_ulong)


class KeyBdInput(ctypes.Structure):
    _fields_ = [("wVk", ctypes.c_ushort),
                ("wScan", ctypes.c_ushort),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class HardwareInput(ctypes.Structure):
    _fields_ = [("uMsg", ctypes.c_ulong),
                ("wParamL", ctypes.c_short),
                ("wParamH", ctypes.c_ushort)]


class MouseInput(ctypes.Structure):
    _fields_ = [("dx", ctypes.c_long),
                ("dy", ctypes.c_long),
                ("mouseData", ctypes.c_ulong),
                ("dwFlags", ctypes.c_ulong),
                ("time", ctypes.c_ulong),
                ("dwExtraInfo", PUL)]


class Input_I(ctypes.Union):
    _fields_ = [("ki", KeyBdInput),
                ("mi", MouseInput),
                ("hi", HardwareInput)]


class Input(ctypes.Structure):
    _fields_ = [("type", ctypes.c_ulong),
                ("ii", Input_I)]


# Actuals Functions

#######################################################################
#
# DirectInput keyboard scan codes
# Taken from http://www.ionicwind.com/guides/emergence/appendix_a.htm
#
#######################################################################
def SetKeyboardConsts(panel_key):
    with open('keybindings.json') as data_file:
        key_binds = json.load(data_file)
        # print(key_binds["keybindings"][panel_key])
        return int(key_binds["keybindings"][panel_key], 16)


def PressKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def ReleaseKey(hexKeyCode):
    extra = ctypes.c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput(0, hexKeyCode, 0x0008 | 0x0002, 0, ctypes.pointer(extra))
    x = Input(ctypes.c_ulong(1), ii_)
    ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))


def KeyStroke(hexKeyCode):
    PressKey(hexKeyCode)
    time.sleep(0.1)
    ReleaseKey(hexKeyCode)
    time.sleep(0.1)
