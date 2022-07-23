#!/usr/bin/python3

import argparse
import sys
import os.path
import zipfile
import enum
from lxml import etree

class Colours():
    DEFAULT = '\033[0m'
    RED = '\033[0;31m'
    GREEN = '\033[0;32m'
    ORANGE = '\033[0;33m'
    BLUE = '\033[0;34m'
    PURPLE = '\033[0;35m'
    CYAN = '\033[0;36m'
    YELLOW = '\033[1;33m'

class FileType(enum.Enum):
    WORD = 1
    POWERPOINT = 2
    EXCEL = 3

FileTypeData = {
    FileType.WORD: {
        'name': 'Microsoft Word',
        'extension': '.docx',
        'path': 'word/',
        'xml': 'settings.xml',
        'nodes': ['writeProtection', 'documentProtection'],
    },
    FileType.POWERPOINT: {
        'name': 'Microsoft Powerpoint',
        'extension': '.pptx',
        'path': 'ppt/',
        'xml': 'presentation.xml',
        'nodes': ['modifyVerifier'],
    },
    FileType.EXCEL: {
        'name': 'Microsoft Excel',
        'extension': '.xlsx',
        'path': 'xl/',
        'xml': 'workbook.xml',
        'nodes': ['fileSharing', 'workbookProtection'],
    },
}

def success(message):
    print(Colours.GREEN + '[+] ' + Colours.DEFAULT + message)

def info(message):
    print(Colours.BLUE + '[#] ' + Colours.DEFAULT + message)

def warning(message):
    print(Colours.ORANGE + '[!] ' + Colours.DEFAULT + message)

def error(message):
    print(Colours.RED + '[-] ' + Colours.DEFAULT + message)

def fatal(code, message):
    error(message)
    sys.exit(code)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Officer Breaker removes a read-only password restriction from a docx/pptx/xlsx file.')
    parser.add_argument('--file', required=True, help='The target docx/pptx/xlsx file.')
    parser.add_argument('--out', required=False, help='The output filename to be created. Default is {file}-writable.{extension}')
    args = parser.parse_args()

    # File exists validation
    if not os.path.isfile(args.file):
        fatal(1, 'Input file does not exist.')
    
    # Parse file path
    filename = os.path.basename(args.file)
    filenameWithoutExtension, extension = os.path.splitext(filename)
    
    # Set output filename if not provided
    if args.out == None or args.out == '':
        args.out = filenameWithoutExtension + '-writable' + extension

    # Find matching filetype data from extension
    data = None
    for currentFileType in FileTypeData:
        if extension == FileTypeData[currentFileType]['extension']:
            data = FileTypeData[currentFileType]

    # Error if not found
    if data == None:
        fatal(2, 'Input file is not a docx, pptx or xlsx.')
    info('Processing ' + data['name'] + ' file...')
    
    # Check file is not in compound file binary format
    with open(args.file, 'rb') as file:
        magic = file.read(8)
        if magic == b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1': # See https://en.wikipedia.org/wiki/List_of_file_signatures
            fatal(3, 'Compound file binary format not supported. This could be due to a legacy document version or an fully encrypted document.')

    try:
        foundPassword = False

        # Read input zip file
        with zipfile.ZipFile(args.file, 'r') as inZip:
            # Write output zip file
            with zipfile.ZipFile(args.out, 'w') as outZip:

                # Copy input zip file comment to output
                outZip.comment = inZip.comment

                # Iterate input zip file files
                for item in inZip.infolist():

                    # If it's not the target file then copy original contents
                    if item.filename != data['path'] + data['xml']:
                        outZip.writestr(item, inZip.read(item.filename))
                        continue

                    # Otherwise remove the password restriction then write
                    tree = etree.fromstring(inZip.read(item.filename))

                    # Iterate each element and child
                    for elem in tree.iter():
                        for child in elem.iter():
                            # Ignore self
                            if elem.tag == child.tag:
                                continue

                            # Remove child if it matches the target node
                            for searchNode in data['nodes']:
                                if searchNode in child.tag:

                                    foundPassword = True

                                    # Print information
                                    info('Removing password protection')

                                    # Find password algorithm, hash and salt
                                    searchAttributes = ['Algorithm Name', 'Hash', 'Salt']
                                    attributes = {}
                                    for attributeKey in child.attrib:
                                        for searchAttribute in searchAttributes:
                                            if searchAttribute.replace(' ', '').lower() in attributeKey.lower():
                                                attributes[searchAttribute] = child.attrib[attributeKey]
                                    
                                    # Print additional password information
                                    for attributeKey in attributes:
                                        info('\t' + attributeKey + ': ' + attributes[attributeKey])

                                    # Remove element
                                    elem.remove(child)

                    # Write the modified XML document
                    outZip.writestr(item, etree.tostring(tree))

        # If we failed to find password attribute, delete duplicated zip and print message
        if not foundPassword:
            os.remove(args.out)
            fatal(4, 'Failed to find password protection in file.')
    except zipfile.BadZipFile as e:
        fatal(5, 'Failed to extract input file.')
    success('File "' + args.out + '" has read-only password protections removed.')