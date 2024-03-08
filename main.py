import os
import tkinter as tk
from itertools import zip_longest
from tkinter import filedialog, simpledialog, messagebox
import re

import pandas as pd
from docxtpl import DocxTemplate


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def search_for_file_path():
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename()

    print(file_path)

    if len(file_path) > 0:
        return file_path
    else:
        raise ValueError('Bad file path')


def findPatches(inputLog):
    patches = []
    for i in inputLog:
        patches.extend(re.findall('(?<=Patch  )[^: ]*', i))
    return patches


def findPatchesID(inputLog):
    patches = []
    for i in inputLog:
        patches.extend(re.findall('(?<=^Patch Id: )[^$][0-9]+', i))
    return patches


def findDesc(inputLog):
    desc = []
    for i in inputLog:
        desc.extend(re.findall('(?<=Patch description:  ")[^"]*', i))
    return desc


def findDescV2(inputLog):
    desc = []
    for i in inputLog:
        desc.extend(re.findall('(?<=^Patch Description: ).*', i))
    return desc


def openLog(name):
    with open(name) as f:
        lines = f.readlines()
        return lines


def makeCoolTable(patches, desc):
    for row in zip_longest(patches, desc):
        # put empty string instead of `None`
        row = ["" if item is None else item
               for item in row]

        # format every item to the same width
        row = ["  {:14}  ".format(item)
               for item in row]

        # join all items in row using `|` and display row
        print("|".join(row))


def findUniqs(patches,desc):
    seen = set()
    uniqP = []
    uniqD = []
    for p,d in zip(patches, desc):
        if p not in seen:
            uniqP.append(p)
            uniqD.append(d)
            seen.add(p)

    # patches = uniqP
    # desc = uniqD

    return uniqP, uniqD


def makeDoc(patches, desc):
    df = pd.DataFrame(list(zip(patches, desc)),
                      columns=['Patch', 'Patch Description'])
    df.to_excel('generated_report.xlsx', index=False)

    version = simpledialog.askstring("Select Target Version", "Target Version (can be left empty)")

    template = DocxTemplate('patch_Template.docx')
    # Declare template variables

    table_contents = []
    for p, d in zip(patches, desc):
        table_contents.append({
            'Patch': p,
            'Description': d
        })
    context = {
        'table_contents': table_contents,
        'version': version or "Target Version",
        'home_environment': "<<INSERT DB/GI HOME>>"
    }
    try:
        print(template.save('generated_report.docx'))
    except PermissionError:
        messagebox.showerror('Error', "Please close generated_report.docx document!")
        raise ValueError("Please close generated_report.docx document")


def output():
    # Get file path
    file_path_variable = search_for_file_path()
    # Open file for reading
    inputLog = openLog(file_path_variable)

    # Create patches and descriptions list
    patches = findPatches(inputLog)
    if len(patches) == 0:
        patches = findPatchesID(inputLog)

    desc = findDesc(inputLog)
    if len(desc) == 0:
        desc = findDescV2(inputLog)

    patches, desc = findUniqs(patches, desc)

    makeCoolTable(patches, desc)
    makeDoc(patches, desc)


if __name__ == '__main__':
    output()
