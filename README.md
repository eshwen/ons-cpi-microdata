# ons-cpi-microdata

Utilities to parse and concatenate spreadsheets comprising CPI microdata and product descriptions.

## Usage

0) Install the dependencies:

   ```sh
   pip install -r requirements.txt
   ```

1) Parse and clean the existing Excel spreadsheets:

   ```sh
   python parse_spreadsheets.py
   ```

2) If they need concatenating for easier ingest into DAP, also run

   ```sh
   python concat_2017_thru_2020_spreadsheets.py
   ```

Please note that these scripts rely on input files that are official and/or sensitive, so may not be included in this repo. Its existence is mainly for safekeeping of the code and visibility in case it needs to be accessed again.
