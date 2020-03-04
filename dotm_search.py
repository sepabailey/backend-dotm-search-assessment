#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Sean Bailey, Janell Huyck, Koren Niles, Chris Wilson, Julita Marshall"

import os
import zipfile


def grabfile(directories):
    hold_files = []
    for root, directories, files in os.walk(directories):
        for filename in files:
            fullPath = os.path.join(root, filename)
            hold_files.append(fullPath)
    return hold_files


def readfile(file_name, text):
    file = zipfile.ZipFile(file_name)
    track_matches = 0
    with file.open('word/document.xml') as f:
        for line in f:
            txt = line.find(text)
            while txt != -1:
                print(line[txt - 40: txt + 40])
                txt = line.find('$', txt + 1)
                track_matches += 1
    return track_matches


def main():
    run_path = grabfile('./dotm_files')
    file_total = 0
    match_total = 0

    for x in run_path:
        if(x.endswith('.dotm')):
            file_total += 1
            matches = readfile(x, '$')
            if (matches > 0):
                match_total += 1
            else:
                continue
    print(match_total)


if __name__ == '__main__':
    main()
