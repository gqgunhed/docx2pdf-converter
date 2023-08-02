#!/usr/bin/env python

import PySimpleGUI as sg
import docx2pdf

# ------------------ GUI here -------------------
sg.theme('SystemDefault')   # Add a touch of color


def converter_loop(filelist):
    """Converts all given docx files into PDF

    :param list filelist: list of filenames to convert
    """
    print("Konvertiere %s .docx Dateien ..." % len(filelist))
    window.Refresh()
    for file in filelist:
        try:
            docx2pdf.convert(file)
            print("%s konvertiert." % file)
        except AttributeError as err:
            print("Fehler beim Konvertieren von %s\nFehler: %s" % (file, str(err.message)))
        window.Refresh()
    print("Fertig.")



# setup a logwindow and redirect the "print" command
logwindow = sg.Multiline(size=(70,10), font=('Courier', 10),
                         autoscroll=True, reroute_stdout=True, reroute_stderr=True)
print = logwindow.print

layout = [
            [sg.Text('Word .docx Dokumente in PDF umwandeln')],
            [sg.Text('Die PDF Dateien werden im gleichen Verzeichnis wie die .docx gespeichert.')],
            [sg.Input(key='_FILES_'), sg.FilesBrowse(file_types=(("Word .docx", "*.docx"),))],
            [sg.OK('Konvertieren', key='_OK_'), sg.Cancel('Ende')],
            [sg.Text('Log')],
            [logwindow],
         ]

# Create the Window
window = sg.Window('Word 2 PDF Converter', layout)

# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Ende': # if user closes window or clicks cancel
        break
    if event == '_OK_':
        filelist = values['_FILES_'].split(';')
        converter_loop(filelist)
    window.Refresh()

window.close()
