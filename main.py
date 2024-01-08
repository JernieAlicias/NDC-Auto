from flask import Flask
from threading import Thread

app = Flask('')

@app.route('/')
def home():
    return "Script is running!"

def run():
    app.run(host='0.0.0.0',port=8080)

def keep_alive():
    t = Thread(target=run)
    t.start()


from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import gspread
import time
import json
import os
    
# Authenticate with Google Sheets API and Google Keep API
scope = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/spreadsheets']
json_str = os.environ.get('GOOGLE_APPLICATION_CREDENTIALS')
creds_dict = json.loads(json_str)
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
client = gspread.authorize(creds)

keep_alive()


while True:

    # TALK SCHEDULE AUTOMATION (Date and Time adjusted by -8 to accomodate to time zone of GMT+8 12:00 PM)
    if datetime.now().weekday() == 0 and datetime.now().hour == 8 and datetime.now().minute <= 59:

        try:
      
            # Set up the Google Sheets Key and Sheet Names Talk Schedule:        
            talk_spreadsheet_key = '1qD0o0aKEwPUzhnIMFAFOojZ1LGAHfFJSAhCIwd5zhq0'
            talk_sheet_name1     = 'Schedule'
            talk_sheet_name2     = 'Outlines'
          
            # Open the Talk Schedule Sheet:
            TSchedule = client.open_by_key(talk_spreadsheet_key).worksheet(talk_sheet_name1)
            Outlines  = client.open_by_key(talk_spreadsheet_key).worksheet(talk_sheet_name2)
    
            # Retrieve the Outline Number of the Talk in Sunday:
            cell_to_find = TSchedule.cell(3, 4)  # Cell D3
            search_value = cell_to_find.value
    
            # Retrieve the Date of the Talk in Sunday:
            cell_of_date = TSchedule.cell(3, 2)  # Cell B3
            date_value = cell_of_date.value
    
            # Find the cell containing the Outline Number in the Outlines sheet:
            outline_number = Outlines.find(search_value)
            print(f'Found Outline # {search_value} in cell {outline_number}')
    
            # Adds a check mark and the date for the Talk's Outline Number after the Weekend Meeting:
            for i in range(5):
                if  Outlines.cell(outline_number.row, 3+i).value == None:
                    Outlines.update_cell(outline_number.row, 9, date_value)
                    Outlines.update_cell(outline_number.row, 3+i,'✔')
                    TSchedule.delete_rows(3)
                    break      
    
            print('Done Updating Talk Schedule. Sleeping for 30 minutes')
            time.sleep(1800) 

        except (ValueError, AttributeError): 
          
            # Set up the Google Sheets Key and Sheet Names Talk Schedule:        
            talk_spreadsheet_key = '1qD0o0aKEwPUzhnIMFAFOojZ1LGAHfFJSAhCIwd5zhq0'
            talk_sheet_name1     = 'Schedule'
          
            # Open the Talk Schedule Sheet and delete the last talk delivered in the schedule:
            TSchedule = client.open_by_key(talk_spreadsheet_key).worksheet(talk_sheet_name1)
            TSchedule.delete_rows(3)
            print('Done Updating Talk Schedule. Sleeping for 6 minutes')
            time.sleep(1860) 
  

    # CART SCHEDULE AUTOMATION (Date and Time adjusted by -8 to accomodate to time zone of GMT+8 6:00 PM)
    if datetime.now().weekday() == 6 and datetime.now().hour == 11 and datetime.now().minute <= 29:

        # Set up the Google Sheets Key and Sheet Names of Cart Schedule:        
        cart_spreadsheet_key = '1q8iIY_r0xg6fvtycdZVvWZWRhQeh7X6XZbXHR_FKiI8'
        cart_sheet_name      = 'Schedule'
      
        # Open the Cart Schedule Sheet:
        sheet = client.open_by_key(cart_spreadsheet_key).worksheet(cart_sheet_name)

        # Define the cell addresses to update in Cart Schedule:
        cell_addresses  = ['C1'  , 'D1'  , 'E1'  , 'F1'  , 'G1'  , 'H1'  , 'I1'  ]
        cell_addresses2 = ['C47' , 'D47' , 'E47' , 'F47' , 'G47' , 'H47' , 'I47' ]
        
        # Define the row ranges to clear in Cart Schedule:
        row_ranges = ['C3:I44', 'C48:I110']

        # Calculate the new dates for next week:
        new_dates = []
        for i in range(7):
            next_week = datetime.now() + timedelta(days=i+1)
            new_dates.append(next_week.strftime('%a — %b %d'))

        # Update the cells in the Google Sheet:
        for i in range(len(cell_addresses)):
            sheet.update_acell(cell_addresses[i],  new_dates[i])
            sheet.update_acell(cell_addresses2[i], new_dates[i])

        # Clear the specified rows in the Google Sheet:
        for row_range in row_ranges:
            cells_to_clear = sheet.range(row_range)
            for cell in cells_to_clear:
                cell.value = ''
            sheet.update_cells(cells_to_clear)

        print('Done Updating Cart Schedule. Sleeping for 30 minutes')
        time.sleep(1800)

    else:
        
        print(f"Waiting to update at Sunday 7:00 PM & Monday 12:00 AM - Current date & time: {datetime.now()}")
        time.sleep(30)



'''

Google Sheet Custom Formula for highlighting the dates of Midweek and Weekend Meeting:

Simple formula that updates the highlight every Monday 12:00 AM:
=AND(B3>=TODAY()-(WEEKDAY(TODAY(),2)-1),B3<=TODAY()+(7-WEEKDAY(TODAY(),2)))

Complex formula that updates the highlight every Sunday 12:00 PM:
=IF(WEEKDAY(TODAY(),2)<7, IF(OR(B3=TODAY()-WEEKDAY(TODAY(),1)+6, B3=TODAY()-WEEKDAY(TODAY(),1)+8), True, False), IF(TEXT(NOW(),"hh:mm:ss")<=TIMEVALUE("12:00:00"), IF(OR(B3=TODAY()-2, B3=TODAY()), True, False), IF(OR(B3=TODAY()+5, B3=TODAY()+7), True, False)))

'''
