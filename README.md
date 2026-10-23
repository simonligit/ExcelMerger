# ExcelMerger
A tool for comparing differences between two Excel files.

Gemini is mainly used to assist development. Copilot is also used.
My purpose is to learn develop a tool for my own requirement with AI assist.
This project is developped using Visual Studio 2026, C# and DotNet 10.0. It's a WPF application.
The first version was developed from 20260106 to 20260110 in my spare time.

V 1.1.0 (2026-05-05) - Add support for password-protected Excel files.
V 1.1.1 (2026-05-08) - Display file path in main window. Make file name input text editable in Open File Dialog.
V 1.2.0 (2026-10-22) - New column mode: compare files by column (files stacked top/bottom, key rows selectable, horizontal minimap).

## Features
- Compare two Excel files and highlight differences.
- Row mode: align rows side-by-side with key columns.
- Column mode: align columns across two files (stacked top/bottom) with key rows and a horizontal minimap.
- Support for multiple sheets within the Excel files.
- Switch between English and Chinese languages.
- Drag-and-drop for loading Excel files.
- Password-protected files are supported.

## Konwn limitations
- Not all Excel features are supported (e.g., charts, images).
- Only compares cell values, not formulas or formatting.

## Improvements
Please feel free to improve this project by submitting issues or pull requests.
Some ideas for future improvements:
- Improve the performance to compare and display large Excel files.
- Support merged cells and different data types.


