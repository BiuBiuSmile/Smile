
import requests
import json
url="https://data.moenv.gov.tw/api/v2/aqx_p_432?api_key=dddeb7a3-46d1-48c5-b4b2-a81a369bf34f"

data = requests.get(url).text

air = json.loads(data)
allair=air['records']
print(allair)

for item in allair:
    print("縣市:",item["county"])
    print("城市:",item["sitename"])
    print("aqi",item["aqi"])
    print("品質",item["status"])
    