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
        ws = wb.active

        dlp1, dlp2 = self.get_sheet_title(ws)

        differences = []
        for row in range(25, ws.max_row + 1):
            d_val = ws[f"D{row}"].value
            t_val = ws[f"T{row}"].value
            e_val = ws[f"E{row}"].value
            u_val = ws[f"U{row}"].value

            if d_val != t_val:
                differences.append((row, d_val, e_val, t_val, u_val))

        return differences, dlp1, dlp2

    def get_sheet_title(self, ws):
        sheet_title = ws.title.strip()
        parts = sheet_title.split()

        if len(parts) != 2:
            raise ValueError(
                f"Expected exactly two DLPs in sheet title, got: '{sheet_title}'"
            )

        return parts[0], parts[1]

    def create_outlook_draft(self, differences, file_name, dlp1, dlp2) -> None:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.Subject = f"Oper deltas found in: {file_name}"
        mail.To = ""

        if differences:
            html_body = f"""
            <html>
            <body>
            <p>The following deltas were found between <b>{dlp1}</b> and <b>{dlp2}</b>:</p>
            <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Arial; font-size: 12px;">
                <tr style="background-color: #f2f2f2;">
                    <th>Row</th>
                    <th>{dlp1}</th>
                    <th>{dlp2}</th>
                </tr>
            """
            for row, d_val, e_val, t_val, u_val in differences:
                html_body += f"""
                <tr>
                    <td>{row}</td>
                    <td>{d_val if d_val is not None else ''}{' '}{e_val if e_val is not None else ''}</td>
                    <td>{t_val if t_val is not None else ''}{' '}{u_val if u_val is not None else ''}</td>
                </tr>
                """
            html_body += """
            </table>
            </body>
            </html>
            """
        else:
            html_body = f"""<p>No differences were found between <b>{dlp1}</b> and <b>{dlp2}</b>.</p>"""

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

        differences, dlp1, dlp2 = self.compare_columns_and_generate_report(file_path)
        self.create_outlook_draft(differences, Path(file_path).name, dlp1, dlp2)
