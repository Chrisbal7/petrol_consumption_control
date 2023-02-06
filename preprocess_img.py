#! /usr/bin/python3

import pytesseract

import logging
import pprint
import os
import argparse
import tempfile

import cv2
import pdf2image
import numpy


from skimage.transform import hough_line, hough_line_peaks
# from skimage.transform import rotate
from skimage.feature import canny
# from skimage.io import imread
from skimage.color import rgb2gray

# import matplotlib.pyplot as plt
from scipy.stats import mode
# import scipy
from PIL import Image
from subprocess import Popen

logging.basicConfig(level=logging.DEBUG, format='%(levelname)s - %(message)s - %(asctime)s')
# logging.disable()

os.chdir('/home/chrisbal/Downloads')

# Construct the argument parser
parser = argparse.ArgumentParser()
parser.add_argument('-i', '--image', help='Path to input image to be processed by OCR', default=None)
parser.add_argument('-alpha', '--alpha', help='Alpha value, increase the value for brightness', default=1)
parser.add_argument('-beta', '--beta', help='Beta value, decrease the value for darkness', default=-32)
parser.add_argument('-b_adapt', '--adaptive', help='adaptive value for binarization', default=5)
parser.add_argument('-pdf', '--pdf', help='Path to the pdf document', default=None)
args = vars(parser.parse_args())

temp_dir = tempfile.TemporaryDirectory()
filename = '{}.png'.format(os.getpid())
temp_img = '{}.png'.format(os.getpid())


def deskewer(img):
    # os.chdir('/home/chrisbal/Downloads/Test')
    image = rgb2gray(img)
    edges = canny(image)

    tested_angles = numpy.deg2rad(numpy.arange(0.1, 180.0))
    h, theta, d = hough_line(edges, theta=tested_angles)

    accum, angles, dists = hough_line_peaks(h, theta, d)
    most_common_angle = mode(numpy.around(angles, decimals=2,))[0]

    skew_angle = numpy.rad2deg(most_common_angle - numpy.pi / 2)
    angle,  = skew_angle
    logging.debug(angle)
    # rotated = rotate(image, skew_angle, cval=1)
    rotated = img
    if 0 < angle < 45:
        rotated = img.rotate(-angle)

    rotated.save(temp_img)

    return temp_img


def extract_text(path):
    text = image_to_string(preprocess_image(path))
    txt = open('text5.txt', 'a')
    txt.write(text)
    txt.close()

    Popen(['open', 'text4.txt'])


def divide_img(path):
    # Load image, grayscale, adaptive threshold
    img1 = Image.open(path)
    image = cv2.imread(path)
    img = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    img_obj = cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 255, 5)

    contours = cv2.findContours(img_obj, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = contours[0] if len(contours) == 2 else contours[1]
    for c in contours:
        cv2.drawContours(img_obj, [c], -1, (255, 0, 0), -1)

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (9, 9))
    opening = cv2.morphologyEx(img_obj, cv2.MORPH_OPEN, kernel, iterations=4)

    contours = cv2.findContours(opening, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    contours = contours[0] if len(contours) == 2 else contours[1]

    i = 0
    img_copy = img1.copy()

    for c in contours:
        x, y, width, height = cv2.boundingRect(c)
        cv2.rectangle(image, (x, y), (x + width, y + height), (36, 255, 12), 3)
        blank = Image.new(mode="RGBA", size=(600, 600), color='white')
        img_cropped = img_copy.crop((x, y, x + width, y + height))
        blank.paste(img_cropped, (0, 0))
        blank.save(f'cropped{i + 1}.png')
        extract_text(f'cropped{i + 1}.png')
        i += 1
    cv2.imwrite('image1.png', image)
    return contours


def preprocess_image(path):
    logging.debug(path)
    image = Image.open(path)
    image.save(filename, dpi=(600, 600))  # Scale

    # img = Image.open(filename)
    alpha = args['alpha']
    beta = args['beta']
    img = cv2.imread(filename)
    adjusted = cv2.convertScaleAbs(img, alpha=alpha, beta=beta)  # contrast

    gray = cv2.cvtColor(adjusted, cv2.COLOR_BGR2GRAY)

    noise_less = cv2.fastNlMeansDenoising(gray, None, 8, 7, 21)  # noise removal
    # Binarizing
    a = args['adaptive']
    img_obj = cv2.adaptiveThreshold(noise_less, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 255, a)
    cv2.imwrite(temp_img, img_obj)

    return temp_img


def image_to_string(img):
    # img = Image.open(path)
    pytesseract.pytesseract.tesseract_cmd = r'/usr/bin/tesseract'
    text = pytesseract.image_to_string(img, lang='fra')
    return text


def close_temp():
    temp_dir.cleanup()
    os.remove(filename)
    os.remove(temp_img)


def extract():
    txt = open('text3.txt', 'w')
    text = ''
    if args['pdf'] is not None:
        text = pdf_to_image(args['pdf'])
    elif args['image'] is not None:
        text = image_to_string(preprocess_image(deskewer(args['image'])))
    txt.write(pprint.pformat(text))
    txt.write(f'\n')
    txt.close()
    Popen(['open', 'text3.txt'])


def pdf_to_image(path):
    text = ''
    images = pdf2image.convert_from_path(path)
    print(type(images))
    i = 0
    for img in images:
        print(type(img))
        img.save(f'page{i + 1}.png')
        image = Image.open(f'page{i + 1}.png')
        text += '\n'
        # text += image_to_string(preprocess_image(deskewer(image)))
        divide_img(f'page{i + 1}.png')
        i += 1

    return text


# divide_img(deskewer(Image.open(args['image'])))
pdf_to_image(args['pdf'])

# extract()
