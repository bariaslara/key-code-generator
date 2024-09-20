# key-code-generator
A Python-based key code generator that logs unique codes with booking IDs to a text file and Excel workbook.

# About This Program

This is a Key Code Generator program designed to create unique 4-digit key codes. The generated codes are saved to a text file to ensure that no duplicates are created. The program also logs the key codes along with their creation date and a user-provided booking ID to an Excel workbook, specifically in the second sheet of the workbook.

## Features:
- Generates a unique 4-digit key code.
- Ensures no duplicate codes are created by checking the saved key codes.
- Logs key codes, creation date, and booking ID in both a text file and an Excel workbook.
- Simple Excel interface to input commands and booking IDs.
- Can be easily run through Excel with a macro button.

## Usage Instructions:
1. Open the Excel workbook that includes this program.
2. Input the booking ID in the provided cell (B2).
3. Use the designated macro button to run the script.
4. The key code, date, and booking ID will be stored in the second sheet of the workbook and a text file.

This program is written in Python and uses the `xlwings` library to interface with Excel.
