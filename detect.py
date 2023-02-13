'''
detect a image in screen
'''

import os
import sys
import time
import cv2 as cv
import numpy as np
import pyautogui
from PIL import Image
from PIL import ImageGrab
from PIL import ImageChops
from PIL import ImageFilter
from PIL import ImageEnhance
from PIL import ImageOps
from PIL import ImageStat
from PIL import ImageDraw
from PIL import ImageFont
from PIL import ImageColor

def get_screen():
    screen = ImageGrab.grab()
    return screen

def get_screen_size():
    screen = get_screen()
    return screen.size

def get_screen_width():
    screen = get_screen()
    return screen.size[0]

def get_screen_height():
    screen = get_screen()
    return screen.size[1]   

def get_screen_pixel(x, y):
    screen = get_screen()
    return screen.getpixel((x, y))


def get_screen_pixel_rgb(x, y):
    screen = get_screen()
    return screen.getpixel((x, y))

def get_screen_pixel_hsv(x, y):
    screen = get_screen()
    return screen.getpixel((x, y))

def get_screen_pixel_hsl(x, y):
    screen = get_screen()
    return screen.getpixel((x, y))

def get_screen_pixel_cmyk(x, y):
    screen = get_screen()
    return screen.getpixel((x, y))

def get_screen_pixel_lab(x, y):
    screen = get_screen()
    return screen.getpixel((x, y))

def get_screen_pixel_yiq(x, y):
    screen = get_screen()
    return screen.getpixel((x, y))

def get_screen_pixel_ycbcr(x, y):
    screen = get_screen()
    return screen.getpixel((x, y))

# check if the image is in the screen
def is_image_in_screen(image, threshold=0.8):
    screen = get_screen()
    result = cv.matchTemplate(np.array(screen), np.array(image), cv.TM_CCOEFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv.minMaxLoc(result)
    if max_val >= threshold:
        return True
    else:
        return False




# capture screen as image   
def capture_screen_as_image():
    screen = get_screen()
    return screen

