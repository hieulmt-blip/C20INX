import os
import requests
import gspread
import datetime
from oauth2client.service_account import ServiceAccountCredentials
from gspread.utils import rowcol_to_a1
import json

# ===== CONFIG FROM ENV =====
API_KEY = os.environ.get("CMC_API_KEY")
SHEET_URL = os.environ.get("SHEET_URL")
SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_CREDENTIALS")
COIN_HEADER_ROW = 6
MARKETCAP_ROW = 7
TIMESTAMP_CELL = "A1"

if not API_KEY:
    raise ValueError("Missing CMC_API_KEY")

if not SERVICE_ACCOUNT_JSON:
    raise ValueError("Missing GOOGLE_CREDENTIALS")

# ===== LOAD GOOGLE CREDS FROM ENV =====
creds_dict = json.loads(SERVICE_ACCOUNT_JSON)

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_dict(
    creds_dict, scope
)

client = gspread.authorize(creds)
sheet = client.open_by_url(SHEET_URL).sheet1


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

    response = requests.get(url, headers=headers, params=params, timeout=15)
    data = response.json()

    if "data" not in data:
        print("API error:", data)
        return

    marketcap_values = []

    for coin in header:
        if coin.strip() == "":
            marketcap_values.append("")
            continue

        try:
            market_cap = float(data["data"][coin]["quote"]["USD"]["market_cap"])
            marketcap_values.append(market_cap)
        except Exception:
            marketcap_values.append("ERROR")

    start_cell = rowcol_to_a1(MARKETCAP_ROW, 1)
    end_cell = rowcol_to_a1(MARKETCAP_ROW, len(marketcap_values))
    range_string = f"{start_cell}:{end_cell}"

    sheet.update([marketcap_values], range_name=range_string)

    now = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

    sheet.update(
        [[f"Last update: {now}"]],
        range_name=TIMESTAMP_CELL
    )

    print("Updated at", now)


if __name__ == "__main__":
    update_marketcap()
