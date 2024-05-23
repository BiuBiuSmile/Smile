# -*- coding: utf-8 -*-
"""
Created on Thu May 23 20:02:50 2024

@author: USER
"""

import requests
import json
url="https://data.ntpc.gov.tw/api/datasets/010e5b15-3823-4b20-b401-b1cf000550c5/json?size=100"
data = requests.get(url).text
#print(data)
bike = json.loads(data)
for item in bike:
    print("站名:",item['sna'])
    print("可借:",item['sbi'])
    print('可停:',item['bemp'])
    print()