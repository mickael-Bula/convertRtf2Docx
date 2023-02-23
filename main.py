#!/bin/py
import re
import os
import win32com.client as win32
from win32com.client import constants


def change_word_format(file_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(file_path)
    doc.Activate()

    # Rename path with .docx
    new_file_abs = os.path.abspath(file_path)
    new_file_abs = re.sub(r'\.\w+$', '.doc', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(
        new_file_abs,
        FileFormat=constants.wdFormatDocument
    )
    doc.Close(False)


# convert rtf to docx and embed all pictures in the final document
def ConvertRtfToDocx(file_path):
    word = win32.Dispatch("Word.Application")
    wdFormatDocumentDefault = 16
    wdHeaderFooterPrimary = 1
    doc = word.Documents.Open(file_path)
    # for pic in doc.InlineShapes:
    #     pic.LinkFormat.SavePictureWithDocument = True
    # for hPic in doc.sections(1).headers(wdHeaderFooterPrimary).Range.InlineShapes:
    #     hPic.LinkFormat.SavePictureWithDocument = True
    doc.SaveAs(
        "C:/Users/bulam/PycharmProjects/pythonProject/mise_en_demeure_heritiers2.docx",
        FileFormat=wdFormatDocumentDefault)
    doc.Close()
    word.Quit()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # change_word_format('C:/Users/bulam/PycharmProjects/pythonProject/Test.rtf')
    # change_word_format('C:/Users/bulam/PycharmProjects/pythonProject/mise_en_demeure_heritiers.rtf')
    ConvertRtfToDocx('C:/Users/bulam/PycharmProjects/pythonProject/mise_en_demeure_heritiers.rtf')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
