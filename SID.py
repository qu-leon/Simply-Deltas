"""
Main script for Simple Deltas
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


def select_excel():
    layout = [
        [
            sg.Text(
                "Select Experiment Compare file to show delta summary.",
                font=("Arial", 11),
            )
        ],
        [
            sg.Text("File location: ", size=(10, 1), font=("Arial", 11)),
            sg.Input(key="-FILE-"),
            sg.FileBrowse("Browse", file_types=(("Excel Files", "*.xlsx;*.xlsm"),)),
        ],
        [sg.Submit(), sg.Button("Cancel")],
    ]

    window = sg.Window(
        "Data Input",
        layout,
        auto_size_text=True,
        icon=resource_path("icons/compare.ico"),
    )
    window.set_icon(resource_path("icons/compare.ico"))
    file_path = None

    while True:
        event, values = window.read()
        if event == "Submit":
            file_path = values["-FILE-"]
            if not file_path:
                sg.popup(
                    "No file selected. Please rerun Simple Deltas.",
                    title="Info",
                    icon=resource_path("icons/compare.ico"),
                )
                break
            break
        elif event in (sg.WIN_CLOSED, "Cancel"):
            break

    window.close()
    return file_path


def main():
    file_path = select_excel()
    if file_path:
        compare = ExcelCompare(file_path)
        compare.run(file_path)

        sg.popup_timed(
            "Deltas generated in Outlook message.",
            title="Finished",
            auto_close_duration=8,
            keep_on_top=True,
            icon=resource_path("icons/compare.ico"),
        )
    else:
        return


if __name__ == "__main__":
    main()
