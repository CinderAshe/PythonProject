import tkinter as tk
import requests
import openpyxl

def search_and_save_to_excel():
    user_input = entry.get()
    
    # Make a request to the card endpoint with the user input
    url = f"https://api.scryfall.com/cards/search/?q={user_input}"
    response = requests.get(url)
    
    if response.status_code == 200:
        data = response.json()
        
        # Save the data to an Excel file
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "Search Results"
        
        # Write headers to the Excel file
        headers = ["Card Name", "Color", "Mana Cost", "Type Line"]
        sheet.append(headers)
        
        # Extract and record specific fields into the Excel file
        for item in data:
            card_name = item.get("name", "")
            colors = ", ".join(item.get("colors", [])) if isinstance(item.get("colors"), list) else ""
            mana_cost = item.get("mana_cost", "")
            type_line = item.get("type_line", "")
            row_data = [card_name, colors, mana_cost, type_line]
            sheet.append(row_data)
        
        wb.save("search_results.xlsx")
        print("Search results saved to search_results.xlsx")
    else:
        print("Failed to retrieve search results from the card endpoint")

# Create the main window
root = tk.Tk()
root.title("Search Cards and Save to Excel")

# Create a label
label = tk.Label(root, text="Please enter the card name to search:")
label.pack()

# Create an entry widget for user input
entry = tk.Entry(root)
entry.pack()

# Create a button to search the cards endpoint and save to Excel
search_button = tk.Button(root, text="Search Cards and Save to Excel", command=search_and_save_to_excel)
search_button.pack()

# Start the main event loop
root.mainloop()