#! usr/bin/python3

""" ---- Renamer -----
 That is a program using to rename data collected in jpg and pdf"""

import os
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('-p', '--path', default='.', help='The path to the test folder')
args = vars(parser.parse_args())

allFiles = os.listdir(args['path'])
os.chdir(args['path'])
i = 0
j = 0

for file in allFiles:
    if file.endswith('jpg'):
        i += 1
        os.rename(fr'./{file}', str(i) + '.jpg')
    if file.endswith('pdf'):
        j += 1
        os.rename(fr'./{file}', str(j) + '.pdf')

