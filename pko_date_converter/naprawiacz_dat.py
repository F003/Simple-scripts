import pandas as pd
from tkinter import Tk, filedialog, messagebox

def convert_date_format(file_path):
    # Read the Excel file
    df = pd.read_excel(file_path)

    # Convert the date columns to datetime format
    df['Data operacji'] = pd.to_datetime(df['Data operacji'], format="'%Y-%m-%d'")
    df['Data waluty'] = pd.to_datetime(df['Data waluty'], format="'%Y-%m-%d'")

    # Convert the date format to dd.mm.yyyy
    df['Data operacji'] = df['Data operacji'].dt.strftime('%d.%m.%Y')
    df['Data waluty'] = df['Data waluty'].dt.strftime('%d.%m.%Y')

    return df

def choose_file():
    # Open a file dialog to select the Excel file
    root = Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx")])
    root.destroy()

    return file_path

def choose_save_location():
    # Open a file dialog to choose the save location for the new Excel file
    root = Tk()
    root.withdraw()
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    root.destroy()

    return save_path

def start_conversion():
    try:
        file_path = choose_file()

        if not file_path:
            messagebox.showinfo("Błąd", "Nie wybrano pliku.")
            return

        df = convert_date_format(file_path)

        save_path = choose_save_location()

        if not save_path:
            messagebox.showinfo("Błąd", "Nie wybrano lokalizacji zapisu.")
            return

        # Save the converted DataFrame as a new Excel file
        df.to_excel(save_path, index=False)
        messagebox.showinfo("Plik zmieniony i zapisany")

    except Exception as e:
        messagebox.showinfo("Błąd", f"Kod: {str(e)}")

# Run the program
start_conversion()
