# Hotel Rates

## Overview

This project is designed to fetch and display hotel rates for various hotels in a specific area. The rates are saved in an Excel file with color-coded cells based on the rates.

## Features

- Fetch hotel rates from an API
- Save rates in an Excel file
- Color-code rates based on comparison

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/yourusername/hotel-rates.git
   ```
2. Navigate to the project directory:
   ```sh
   cd hotel-rates
   ```
3. Install dependencies:
   ```sh
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```sh
   python app.py
   ```
2. The rates will be saved in an Excel file named `rates-<today's date>.xlsx`.

## Configuration

- Modify the `start` and `end` variables in `app.py` to change the date range for fetching rates.
- Update the `HOTEL_KEYS` dictionary in `app.py` to add or remove hotels.

## License

This project is licensed under the MIT License.
