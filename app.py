import requests
from datetime import date, timedelta
import json
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side

CLE_AREA_CODE = "g50207"
BASE_URL = f"https://data.xotelo.com/api/rates?currency=USD"

COMFORT_INN_KEY = "d101729"
HOTEL_KEYS = {
    "Holiday Inn": "d122215",
    "Hotel Indigo": "d122210",
    "Drury Plaza": "d9862281",
    "Residence Inn": "d122181",
    "DoubleTree": "d1372556",
    "Hilton Garden": "d240589",
    "Hampton Inn": "d111524",
    "The Westin": "d95215",
    "Crowne Plaza": "d95218",
    "Marriott@Key": "d95183",
    "Aloft": "d4375420"
}

start = 1 #start tomorrow, since api disabled pulling rates for today 
end = 31 #inclusive, run for 31 days

def main():
    today_date = date.today()
    wb, sheet = initialize_workbook()

    intro_string = f'''
    Made with ❤️  by Darshan Thakkar
    This program pulls hotel rates from the Xotelo API for a list of hotels in Cleveland, Ohio.
    These rates are gathered from Trip Advisor, which lists rates from OTA sites like Expedia and Booking.com.
    The rates are pulled for a 30-day period starting tomorrow.
    Rates will be saved to file: rates-{today_date}.xlsx
    '''

    print(intro_string)

    for i in range(start, end+1):
        target_date = today_date + timedelta(days=i)
        formatted_date = target_date.strftime("%m/%d")

        date_cell = sheet.cell(row=1+i, column=1, value=formatted_date)
        date_cell.font = Font(bold=True)
        print(f"Pulling rates for date: {formatted_date}")

        # Comfort Inn
        comfort_inn_rate = get_rate(COMFORT_INN_KEY, target_date)
        comfort_inn_cell = sheet.cell(row=1+i, column=2)
        comfort_inn_cell.alignment = Alignment(horizontal='center')
        if comfort_inn_rate is not None:
            comfort_inn_cell.value = comfort_inn_rate
        else:
            comfort_inn_cell.value = "N/A"

        # dict to store hotel name and rate
        hotel_rates = {}

        # for all hotels in HOTEL_KEYS, get the rate and store in hotel_rates dict
        for hotel_name, hotel_key in HOTEL_KEYS.items():
            rate = get_rate(hotel_key, target_date)
            if rate is not None:
                hotel_rates[hotel_name] = rate

        # sort the dict by rate
        hotel_rates = sorted(hotel_rates.items(), key=lambda x: x[1])

        # write the first three hotel name and rate to the sheet
        for idx, (hotel_name, rate) in enumerate(hotel_rates[:3]):
            hotel_name_cell = sheet.cell(row=1+i, column=3+idx*2, value=hotel_name)
            hotel_rate_cell = sheet.cell(row=1+i, column=4+idx*2, value=rate)
            hotel_rate_cell.alignment = Alignment(horizontal='center')
            

    finalize_sheet(sheet)
    wb.save(f"rates-{today_date}.xlsx")
    
def initialize_workbook():
    wb = Workbook()
    # wb = load_workbook(f"rates-{today_date}.xlsx")
    sheet = wb.active
    write_header(sheet)
    return wb, sheet

def get_rate(hotel_key, date):
    url = f"{BASE_URL}&hotel_key={CLE_AREA_CODE}-{hotel_key}&chk_in={date}&chk_out={date + timedelta(days=1)}"

    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
    except requests.exceptions.RequestException as e:
        print(f"Network error for hotel: {hotel_key} on date: {date} - {e}")
        return None
    except json.JSONDecodeError as e:
        print(f"JSON decoding error for hotel: {hotel_key} on date: {date} - {e}")
        return None

    if "error" in data and data["error"] is not None:
        print(f"Error {data["error"]["status_code"]}: {data["error"]["message"]}")
        return None

    if "result" not in data or data["result"] is None:
        print("Error: Missing 'result' in response.")
        return None

    rates = data["result"]["rates"]
    if not rates:
        print(f"No rates found for hotel: {hotel_key} for date: {date}")
        return None
    
    return rates[0]["rate"]

def write_header(sheet):
    sheet.cell( row=1, column=1 , value="Date")
    
    sheet.cell(row=1, column=2, value = "Comfort Inn")

    sheet.cell(row=1, column=3, value = "Hotel 1")
    sheet.cell(row=1, column=4, value = "Rate 1")
    
    sheet.cell(row=1, column=5, value = "Hotel 2")
    sheet.cell(row=1, column=6, value = "Rate 2")

    sheet.cell(row=1, column=7, value = "Hotel 3")
    sheet.cell(row=1, column=8, value = "Rate 3")

def finalize_sheet(sheet):

    # bold and center the headers row
    for column in range(1, 9):
        cell = sheet.cell(row=1, column=column)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    sheet.column_dimensions['A'].width = 9 # date column
    sheet.column_dimensions['B'].width = 14 # comfort inn rate column
    # hotel name columns
    sheet.column_dimensions['C'].width = 15
    sheet.column_dimensions['E'].width = 15
    sheet.column_dimensions['F'].width = 10
    # hotel rate columns
    sheet.column_dimensions['D'].width = 10
    sheet.column_dimensions['G'].width = 15
    sheet.column_dimensions['H'].width = 10

    # Define a thin border for all sides
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # apply border and font 14 to all cells in the worksheet
    for row in sheet.iter_rows():
        # set row height
        sheet.row_dimensions[row[0].row].height = 21
        
        for cell in row:
            cell.border = thin_border
            cell.font = Font(size=14, bold=cell.font.bold)

    # set custom margins
    sheet.page_margins.left = 0.5
    sheet.page_margins.right = 0.25
    sheet.page_margins.top = 0.75
    sheet.page_margins.bottom = 0.75
    sheet.page_margins.header = 0.3
    sheet.page_margins.footer = 0.3

if __name__ == "__main__":
    main()