import requests
import pandas as pd
import json

# --- CONFIGURATION ---
# Replace these with the IDs of the groups you want to check
GROUP_IDS = [8007841, 32931874, 32046157, 32519643,
             32364971, 34024127, 33379579, 34564916,
             34755399, 34863137, 6305239, 34944955,
             10914487, 35081672, 35094828, 35190717,
             35257340, 35259863, 35326763, 35378650,
             35475162, 35559102, 35615952, 35657397,
             35708128, 35805286, 35805297, 359009073,
             281916316, 35819196, 35909698, 6997309,
             35876921, 35909691, 35849683, 35792179,
             1057322826, 637734737, 914302392, 335772096,
             802097319, 56650360, ] 
TEXT_FILENAME = "games_data.txt"
EXCEL_FILENAME = "Roblox_Group_Stats.xlsx"

def fetch_group_games():
    all_games = []
    
    for group_id in GROUP_IDS:
        print(f"Fetching games for group {group_id}...")
        # Roblox API for group games
        url = f"https://games.roblox.com/v2/groups/{group_id}/games?accessFilter=1&sortOrder=Asc&limit=100"
        
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            for game in data.get('data', []):
                game_info = {
                    "Name": game.get("name"),
                    "Link": f"https://www.roblox.com/games/{game.get('rootPlaceId')}",
                    "Visits": game.get("placeVisits", 0)
                }
                all_games.append(game_info)
        else:
            print(f"Failed to fetch data for group {group_id}")

    # Save to Text File (JSON format for easy reading later)
    with open(TEXT_FILENAME, "w", encoding="utf-8") as f:
        json.dump(all_games, f, indent=4)
    print(f"Saved raw data to {TEXT_FILENAME}")

def convert_to_excel():
    # Read from the text file
    with open(TEXT_FILENAME, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    df = pd.DataFrame(data)
    
    # Calculate Total Visits
    total_visits = df["Visits"].sum()
    
    # Create the Excel file
    with pd.ExcelWriter(EXCEL_FILENAME, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Games')
        
        # Access the workbook to add the total sum cell
        workbook = writer.book
        worksheet = writer.sheets['Games']
        
        # Add a row at the bottom for the Total
        last_row = len(df) + 2
        worksheet.cell(row=last_row, column=2, value="TOTAL VISITS:")
        worksheet.cell(row=last_row, column=3, value=total_visits)

    print(f"Excel file '{EXCEL_FILENAME}' created successfully!")

if __name__ == "__main__":
    fetch_group_games()
    convert_to_excel()