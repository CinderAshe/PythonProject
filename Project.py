import requests
import openpyxl
import os
import tkinter as tk
from tkinter import messagebox

def search_card():
    card_name = entry_card_name.get()
    
    # Search for the card on Scryfall
    url = f"https://api.scryfall.com/cards/named?fuzzy={card_name}"
    response = requests.get(url)
    data = response.json()

    # Extracting relevant information
    name = data.get('name', '')
    colors = ', '.join(data.get('colors', []))
    card_url = data.get('scryfall_uri', '')

    # Check if the Excel file already exists
    file_name = "MyCardLibrary.xlsx"
    if not os.path.exists(file_name):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Card Data"
        ws.append(["Card Name", "Colors", "URL"])
    else:
        wb = openpyxl.load_workbook(file_name)
        ws = wb.active

    # Data
    ws.append([name, colors, card_url])

    # Save the Excel file
    wb.save(file_name)
    messagebox.showinfo("Success", f"Data saved to {file_name}")

# Create GUI
root = tk.Tk()
root.title("Card Search and Save")

label_card_name = tk.Label(root, text="Enter the card name:")
label_card_name.grid(row=0, column=0, padx=10, pady=10)

entry_card_name = tk.Entry(root)
entry_card_name.grid(row=0, column=1, padx=10, pady=10)

btn_search = tk.Button(root, text="Search and Save", command=search_card)
btn_search.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()
