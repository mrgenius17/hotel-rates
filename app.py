import requests
from datetime import date, timedelta
import json
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill

CLE_AREA_CODE = "g50207"
BASE_URL = f"https://data.xotelo.com/api/rates?currency=USD"

HOTEL_KEYS = {
    "Comfort Inn": "d101729",
    "Holiday Inn": "d122215",
    "Hotel Indigo": "d122210",
    "Drury Plaza": "d9862281",
    "Residence Inn": "d122181",
    "DoubleTree": "d1372556",
    "Hilton Garden": "d240589",
    "Hampton Inn": "d111524",
    "The Westin": "d95215",
    "Crowne Plaza": "d95218",
    "Marriott Key": "d95183",
    "Aloft": "d4375420"
}

start = 1 #start tomorrow, since api disabled pulling rates for today 
end = 31 #run for 30 days


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

    for i in range(start, end):
        target_date = today_date + timedelta(days=i)

        formatted_date = target_date.strftime("%m/%d")
        sheet.cell(row=2+i, column=1, value=formatted_date)
        print(f"Pulling rates for date: {formatted_date}")

        for j, hotel_key in enumerate(HOTEL_KEYS.keys(), start=2):
            rate = get_rate(hotel_key, target_date)
            if rate is not None:
                sheet.cell(row=2+i, column=j, value=rate)

    color_rates(sheet)
    wb.save(f"rates-{today_date}.xlsx")
    
def initialize_workbook():
    wb = Workbook()
    # wb = load_workbook(f"rates-{today_date}.xlsx")
    sheet = wb.active
    write_header(sheet)
    return wb, sheet

def get_rate(hotel_key, date):
    url = f"{BASE_URL}&hotel_key={CLE_AREA_CODE}-{HOTEL_KEYS[hotel_key]}&chk_in={date}&chk_out={date + timedelta(days=1)}"

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
    for i, hotel_name in enumerate(list(HOTEL_KEYS.keys()), start=2):
        sheet.cell(row=1, column=i, value=hotel_name)

    sheet.cell(row=2, column=3, value="Our Rate")
    sheet.cell(row=2, column=5, value="Next three higher rates")
    sheet.cell(row=2, column=8, value="Lower than our rate")

    for col_idx in range(1, sheet.max_column + 1):
        col_letter = sheet.cell(row=1, column=col_idx).column_letter
        sheet.column_dimensions[col_letter].width = 12

def color_rates(sheet):
    RED_FILL = PatternFill(start_color="FF6961", end_color="FF6961", fill_type="solid")
    YELLOW_FILL = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")
    GREEN_FILL = PatternFill(start_color="77DD77", end_color="77DD77", fill_type="solid")

    sheet.cell(row=2, column=2).fill = GREEN_FILL
    sheet.cell(row=2, column=4).fill = YELLOW_FILL
    sheet.cell(row=2, column=7).fill = RED_FILL

    for i in range(3, end+2):
        sheet.cell(row=i, column=2).fill = GREEN_FILL
        if sheet.cell(row=i, column=2).value is None:
            continue
        comfortinn_rate = sheet.cell(row=i, column=2).value

        three_lowest_rates = []

        for j in range(3, len(HOTEL_KEYS.keys()) + 2):
            if sheet.cell(row=i, column=j).value is None:
                continue
            rate = sheet.cell(row=i, column=j).value
            if not isinstance(rate, (int)):
                continue
            cell = sheet.cell(row=i, column=j)
            
            if rate < comfortinn_rate:
                sheet.cell(row=i, column=j).fill = RED_FILL
                continue
            
            if len(three_lowest_rates)<3:
                three_lowest_rates.append(cell)
                three_lowest_rates.sort(key=lambda x: x.value)
            elif rate < three_lowest_rates[2].value:
                three_lowest_rates[2] = cell
                three_lowest_rates.sort(key=lambda x: x.value)


        for cell in three_lowest_rates:
            if cell is not None:
                cell.fill = YELLOW_FILL        

if __name__ == "__main__":
    main()