import requests
import openpyxl
import os

def search_card(card_name):
    # Search for the card on Scryfall
    url = f"https://api.scryfall.com/cards/named?fuzzy={card_name}"
    response = requests.get(url)
    data = response.json()

    # Extracting Card Info
    name = data.get('name', '')
    colors = ', '.join(data.get('colors', [])) if data.get('colors') else ''
    card_url = data.get('scryfall_uri', '')

    # Check to Make sure Fine Already Exists
    file_name = "MyCardLibrary.xlsx"
    if not os.path.exists(file_name):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Card Data"
        ws.append(["Card Name", "Colors", "URL", "Count"])
    else:
        wb = openpyxl.load_workbook(file_name)
        ws = wb.active

    # Duplicate Prevention check
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
        if row[0].value == name:
            count = ws.cell(row=row[0].row, column=4).value
            if count is None:
                count = 0
            ws.cell(row=row[0].row, column=4, value=count + 1)
            break
    else:
        ws.append([name, colors, card_url, 1])

    # Save data to excel file
    wb.save(file_name)

def read_file_and_process_searches(file_path):
    """
    Reads a text document line by line and uses each line as input for the search_card function.
    """
    if not os.path.isfile(file_path):
        print(f"File {file_path} does not exist.")
        return

    with open(file_path, 'r') as file:
        lines = file.readlines()
    
    for line in lines:
        line = line.strip()                        # Remove any whitespace
        if line:                                 # Only process non-empty lines
            search_card(line)

# Example usage
if __name__ == "__main__":
    # Path to the text document with card names
    file_path = 'MyMagicCardList.txt'
    
    # Process the searches
    read_file_and_process_searches(file_path)
