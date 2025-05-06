"""
Main script for Simply Deltas
"""

import sys
import os
import warnings
import PySimpleGUI as sg
from Deltas import ExcelCompare

warnings.filterwarnings("ignore")

sg.theme("SystemDefault")


def resource_path(relative_path):
    """Get the absolute path to the resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def main():
    """
    Main GUI function to run
    """
    input_layout = [
        [
            sg.Text(
                "Enter experiment compare file.",
                font=("Arial", 11),
            )
        ],
        [
            sg.Text("File location: ", size=(10, 1), font=("Arial", 11)),
            sg.Input(key="-FILE-"),
            sg.FileBrowse("Browse"),
        ],
        [sg.Submit()],
    ]

    window = sg.Window("Data Input", input_layout, auto_size_text=True)
    window.set_icon(resource_path("icons/book.ico"))
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            window.close()
            sys.exit()
        elif event == "Submit":
            file_location = values["-FILE-"]
            window.close()
            break

    compare = ExcelCompare(file_location)
    compare.run()


if __name__ == "__main__":
    main()
