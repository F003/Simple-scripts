import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Function to convert the date format
def convert_date_format(input_file, output_file):
    df = pd.read_excel(input_file)  # Read the Excel file into a DataFrame
    df['Date Column'] = pd.to_datetime(df['Date Column']).dt.strftime('%d.%m.%Y')
    df.to_excel(output_file, index=False)  # Save the DataFrame to a new Excel file

# Function to handle the button click event
def start_conversion():
    input_path = filedialog.askopenfilename(title='Select Input File', filetypes=[('Excel Files', '*.xls;*.xlsx')])
    output_path = filedialog.asksaveasfilename(title='Select Output File', defaultextension='.xlsx')
    convert_date_format(input_path, output_path)
    print("Conversion completed successfully!")

# Create a simple GUI using tkinter
window = tk.Tk()
window.title("Date Format Conversion")
button = tk.Button(window, text="Start Conversion", command=start_conversion)
button.pack()
window.mainloop()
