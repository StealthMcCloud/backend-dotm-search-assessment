#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "StealthMcCloud"

import os
from zipfile import ZipFile
import argparse

def main(directory, search_text):
    path = directory + "/" if directory != None else "./"
    files = os.listdir(path)
    files_matched = 0
    files_searched = 0
    text_length = len(search_text)
    print "Searching directory " + path + " for input " + search_text
    for searched_file in files:
        if searched_file.endswith(".dotm"):
            files_searched += 1
            with ZipFile(path + searched_file, "r") as zip:
                text = zip.read("word/document.xml")
                if search_text in text:
                    files_matched += 1
                    print "Match found: " + path + searched_file
                    len1 = len(text)
                    for i in range(len1):                           
                         if text[i:i + text_length] == search_text:
                            start_text = 0 if i < 40 else i - 40
                            end_text = len1 if i + 40 + text_length > len1 else i + 40 + text_length
                            print "..." + text[start_text:end_text] + "..."
    print "Total files searched: " + str(files_searched)
    print "Total files matched: " + str(files_matched)

if __name__ == '__main__':
    """Searches input of directory of dotm files for text"""
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir",help="directory to search")
    parser.add_argument("search_text", help="the text to search for")
    args = parser.parse_args()
    main(args.dir,args.search_text)