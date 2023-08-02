#!/usr/bin/env python

import PySimpleGUI as sg
import docx2pdf

# ------------------ GUI here -------------------
sg.theme('SystemDefault')   # Add a touch of color


def converter_loop(filelist):
    """Converts all given docx files into PDF

    :param list filelist: list of filenames to convert
    """
    print("Converting %s .docx file(s) ..." % len(filelist))
    window.Refresh()
    for file in filelist:
        try:
            docx2pdf.convert(file)
            print("%s converted." % file)
        except AttributeError as err:
            print("Error converting %s\nError: %s" % (file, str(err.message)))
        window.Refresh()
    print("Fertig.")



# setup a logwindow and redirect the "print" command
logwindow = sg.Multiline(size=(70,10), font=('Courier', 10),
                         autoscroll=True, reroute_stdout=True, reroute_stderr=True)
print = logwindow.print

layout = [
            [sg.Text('Convert Word .docx documents to PDF')],
            [sg.Text('PDF files will be saved to the same directory as the .docx.')],
            [sg.Input(key='_FILES_'), sg.FilesBrowse(file_types=(("Word .docx", "*.docx"),))],
            [sg.OK('Convert', key='_OK_'), sg.Cancel('Quit')],
            [sg.Text('Log')],
            [logwindow],
         ]

# Create the Window
window = sg.Window('Word 2 PDF Converter', layout)

# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Quit': # if user closes window or clicks cancel
        break
    if event == '_OK_':
        filelist = values['_FILES_'].split(';')
        converter_loop(filelist)
    window.Refresh()

window.close()
