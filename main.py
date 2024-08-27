from ExcelProcessor import ExcelProcessor
import tkinter as tk
from tkinter import filedialog
import os

# TO-DO
# -Currency
# -exhange rate
initial_dir = os.path.normpath("C:/Users/MikePentland/Polio/PBC/data")

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    # Directory selection prompt
    data_dir = filedialog.askdirectory(title="Select campaign data DIRECTORY", initialdir=initial_dir)

    xl_processor = ExcelProcessor(data_dir)

    # Run cleaning operation loop on every xlsx in directory
    for file in xl_processor.xl_files:
        xl_processor.process_file(file)
    xl_processor.concat_xlsx_files()

    