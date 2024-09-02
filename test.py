import requests
from datetime import datetime, timedelta
from tqdm import tqdm

url = "https://api.coindcx.com/exchange/v1/derivatives/futures/data/active_instruments"

response = requests.get(url)
data = response.json()

x = []
now = datetime.now()
from_day = now - timedelta(days=1)
for _, pair in enumerate(tqdm(data)):
    url = "https://public.coindcx.com/market_data/candlesticks"
    query_params = {
        "pair": pair,
        "from": int(from_day.timestamp()),
        "to": int(now.timestamp()),
        "resolution": "15",
        "pcode": "f"
    }
    response = requests.get(url, params=query_params)
    data = response.json()
    if data['data'][-1]['close'] < 100 and data['data'][-1]['volume'] > 2000000:
        x.append((pair, data['data'][-1]['volume']))

def s(d):
    return d[1]

x.sort(key=s)
print(len(x), x)