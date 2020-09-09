#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "???"
import argparse
import sys
import os
import zipfile


def get_file_dict(text, path):
    """Return a dictionry with keys being the path
     and value being the text snippet from dotm files with a specified search text """
    file_dict = {}
    file_list = os.listdir(path)
    for dotm_file in file_list:
        full_path = os.path.join(path, dotm_file)
        if os.path.splitext(full_path)[1] == '.dotm':
            extracted_folder = "new_filez"
            with zipfile.ZipFile(full_path, 'r') as myzip:
                myzip.extractall(extracted_folder)
                for dirpath, dirnames, filenames in os.walk(extracted_folder):
                    for filename in filenames:
                        if filename == 'document.xml':
                            with open(os.path.join(dirpath, filename)) as f:
                                output = f.read()
                            if output.count(text):
                                file_dict[os.path.join(path, dotm_file)] =\
                                    (output[output.index(
                                        text) - 40:output.index(text) + 40])
    print(file_dict)
    return file_dict


def create_parser():
    """Returns an ArgumentParser that takes a single positional argument and an optional --dir flag"""
    parser = argparse.ArgumentParser()
    parser.add_argument('search_text', help="Text to search for")
    parser.add_argument('--dir', help="Path to search in")
    return parser


def main(args):
    """
    Main creates a parser and looks at the namespace if there is a --dir flag called
    And prints the desired format of the solution
    """
    parser = create_parser()
    ns = parser.parse_args(args)
    path = os.getcwd()
    if ns.dir:
        path = os.path.abspath(ns.dir)
    file_dict = get_file_dict(ns.search_text, path)
    print(f'Searching directory {path} for text \'{ns.search_text}\'')
    for file_name, file_line in file_dict.items():
        print(f'Match found in file {file_name}\n...{file_line}...')
        file_len = len([f for f in os.listdir(path) if f.endswith('dotm')])
    print(f'Total dotm files searched: {file_len}')
    print(f'Total dotm files matched: {len(file_dict)}')


if __name__ == '__main__':
    main(sys.argv[1:])
