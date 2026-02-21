import os
import requests
import gspread
import datetime
import json
import time
from oauth2client.service_account import ServiceAccountCredentials
from gspread.utils import rowcol_to_a1
from datetime import timezone, timedelta

# =========================
# ENV CONFIG
# =========================

API_KEY = os.environ.get("CMC_API_KEY")
SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_CREDENTIALS")

SPREADSHEET_ID = "1rPXYy-_zjwgJKBPUFMXoVGb0y8dI5G7yt8QZZtTX0uU"

COIN_HEADER_ROW = 6
MARKETCAP_ROW = 7
TIMESTAMP_CELL = "A1"

if not API_KEY:
    raise ValueError("Missing CMC_API_KEY")

if not SERVICE_ACCOUNT_JSON:
    raise ValueError("Missing GOOGLE_CREDENTIALS")

# =========================
# GOOGLE AUTH
# =========================

creds_dict = json.loads(SERVICE_ACCOUNT_JSON)

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_dict(
    creds_dict,
    scope
)

client = gspread.authorize(creds)
sheet = client.open_by_key(SPREADSHEET_ID).sheet1

VN_TZ = timezone(timedelta(hours=7))


# =========================
# UPDATE FUNCTION
# =========================

def update_marketcap():
    print("Updating...")

    header = sheet.row_values(COIN_HEADER_ROW)
    coins = [c.strip() for c in header if c.strip() != ""]

    if not coins:
        print("No coins found.")
        return

    url = "https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest"

    headers = {
        "Accepts": "application/json",
        "X-CMC_PRO_API_KEY": API_KEY
    }

    params = {
        "symbol": ",".join(coins),
        "convert": "USD"
    }

    response = requests.get(url, headers=headers, params=params, timeout=20)
    data = response.json()

    if "data" not in data:
        print("CMC error:", data)
        return

    marketcap_values = []

    for coin in header:
        if coin.strip() == "":
            marketcap_values.append("")
            continue

        try:
            market_cap = float(
                data["data"][coin]["quote"]["USD"]["market_cap"]
            )
            marketcap_values.append(market_cap)
        except:
            marketcap_values.append("ERROR")

    start_cell = rowcol_to_a1(MARKETCAP_ROW, 1)
    end_cell = rowcol_to_a1(MARKETCAP_ROW, len(marketcap_values))
    range_string = f"{start_cell}:{end_cell}"

    sheet.update(values=[marketcap_values], range_name=range_string)

    now = datetime.datetime.now(VN_TZ).strftime("%Y-%m-%d %H:%M:%S GMT+7")

    sheet.update(
        values=[[f"Last update: {now}"]],
        range_name=TIMESTAMP_CELL
    )

    print("Updated at", now)


# =========================
# EXACT 5-MIN LOOP
# =========================

def sleep_until_next_5_min():
    now = datetime.datetime.now(timezone.utc)
    seconds = (5 - now.minute % 5) * 60 - now.second
    if seconds <= 0:
        seconds += 300
    print(f"Sleeping {seconds}s...")
    time.sleep(seconds)


# =========================
# MAIN LOOP
# =========================

if __name__ == "__main__":
    print("Worker started...")
    while True:
        sleep_until_next_5_min()
        try:
            update_marketcap()
        except Exception as e:
            print("Error:", e)
