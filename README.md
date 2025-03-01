# Docx Tag Processor

## Overview

The Docx Tag Processor is a simple tool designed to automate the process of replacing tagged placeholders within DOCX files. This program scans a folder containing DOCX templates, identifies the tags within those files, and allows the user to input replacement values for the tags. The processed files are then saved to an output folder with the tags replaced by the user-provided values.

## Features

- **Scan DOCX Files**: Automatically scans the selected folder for DOCX files and identifies tagged placeholders.
- **User Input**: Prompts the user to input replacement values for each identified tag.
- **Batch Processing**: Processes all DOCX files in the folder and saves the modified files to an output folder.
- **Graphical Interface**: Provides a simple and user-friendly graphical interface using Tkinter.

## How to Use

1. **Open the Program**: Run the program by executing the Python script.
2. **Select Folder**: Click the "Open Folder" button to select the folder containing the "Template Files" folder with DOCX templates.
3. **Enter Replacement Values**: The program will display input fields for each identified tag. Enter the replacement values for each tag.
4. **Process Files**: Click the "Process Files" button to replace the tags in all DOCX files and save the modified files to the "output" folder within the selected directory.
5. **Success**: A message box will inform you when the files have been successfully processed.

## Prerequisites

- Python 3.x
- Required Python packages: `python-docx`, `tkinter`

## Installation

To install the required Python packages, run the following command:
```sh
pip install python-docx
