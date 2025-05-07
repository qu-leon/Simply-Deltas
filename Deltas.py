"""
Library functions to manipulate EMS flow compare excel files
"""

import sys
import os
import openpyxl
import win32com.client as win32
from pathlib import Path
import PySimpleGUI as sg

sg.theme("SystemDefault")


def resource_path(relative_path):
    """Get the absolute path to the resource, works for dev and for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class ExcelCompare:
    """Class functions for handling EMS flow compare"""

    def __init__(self, save_file) -> None:
        self.save_file = save_file

    def compare_columns_and_generate_report(self, file_path) -> None:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active  # Adjust if needed for specific sheet

        differences = []
        for row in range(25, ws.max_row + 1):
            d_val = ws[f"D{row}"].value
            t_val = ws[f"T{row}"].value

            if d_val != t_val:
                differences.append((row, d_val, t_val))

        return differences

    def create_outlook_draft(self, differences, file_name) -> None:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.Subject = f"Oper deltas found in: {file_name}"
        mail.To = ""

        if differences:
            html_body = """
            <html>
            <body>
            <p>The following deltas were found <b>Lot A</b> and <b>Lot B</b>:</p>
            <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Arial; font-size: 12px;">
                <tr style="background-color: #f2f2f2;">
                    <th>Row</th>
                    <th>Lot A</th>
                    <th>Lot B</th>
                </tr>
            """
            for row, d_val, t_val in differences:
                html_body += f"""
                <tr>
                    <td>{row}</td>
                    <td>{d_val if d_val is not None else ''}</td>
                    <td>{t_val if t_val is not None else ''}</td>
                </tr>
                """
            html_body += """
            </table>
            </body>
            </html>
            """
        else:
            html_body = "<p>No differences were found between <b>Lot A</b> and <b>Lot B</b>.</p>"

        mail.HTMLBody = html_body
        mail.Display()

    def validate_excel_format(self, ws) -> None:
        errors = []

        # Check if columns D and T are within range
        if ws.max_column < 20:
            errors.append("The sheet appears to be missing Lotplan opers.")

        # Check if rows from 25 exist
        if ws.max_row < 25:
            errors.append("The sheet does not have 25 or more rows.")

        return errors

    def run(self, file_path) -> None:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active

        format_errors = self.validate_excel_format(ws)
        if format_errors:
            sg.popup_error(
                "File validation failed:\n\n" + "\n".join(format_errors),
                title="Format Error",
                icon=resource_path("icons/compare.ico"),
            )
            return

        differences = self.compare_columns_and_generate_report(file_path)
        self.create_outlook_draft(differences, Path(file_path).name)
