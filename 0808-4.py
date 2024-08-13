# -*- coding: utf-8 -*-
"""
Created on Thu Aug  8 21:23:42 2024

@author: USER
"""

from bs4 import BeautifulSoup

import requests

url ="https://news.tvbs.com.tw/realtime"

header = {
    
'user-agent':
'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36',
'cookie':
'_gid=GA1.3.683060117.1723123434; cho_weather=%E8%87%BA%E5%8C%97%E5%B8%82; FPID=FPID2.3.zBfcr9CgwxJSzqyhFMska67emRimAC35rxPetEsNaJo%3D.1723123434; FPLC=7VEJEqGYsEULHl%2FAIegAKlV96MVGKn4tV1yBSkwab%2BiQFW3FpkI78EyUxw9aXZYIaJdHXpdbBPNWH0oF08U6rJJfYr0NRgbs8H26miQdA5enEGQVDANbwHS%2F1jXT%2FA%3D%3D; AMP_TOKEN=%24NOT_FOUND; _fbp=fb.2.1723123453642.35135002677764984; _clck=4ssbkv%7C2%7Cfo5%7C0%7C1681; _cc_id=c3c2ec5bfcc9e854b615c5545a14d2cd; panoramaId_expiry=1723728254031; panoramaId=8eb9e09c0d518a35b2044861ec3c16d539381239d2d67f260882c8bddafb3a71; panoramaIdType=panoIndiv; truvid_protected={"val":"c","level":2,"geo":"TW","timestamp":1723123458}; trc_cookie_storage=taboola%2520global%253Auser-id%3D38a96bce-9f00-49a5-ab58-3d2f7b4d62ca-tuctc2a5b13; _gat=1; _ga_F0SK2CNW1N=GS1.1.1723123434.1.1.1723123749.60.0.0; __gads=ID=da5e337253306232:T=1723123453:RT=1723123753:S=ALNI_MakobbUXX44rsuMfEJGQwWaADGFwQ; __gpi=UID=00000eb940a62ae8:T=1723123453:RT=1723123753:S=ALNI_MZRWcnAv2c72QCEjQBZXPDsqo8ZTQ; __eoi=ID=5496048f32c2d30e:T=1723123453:RT=1723123753:S=AA-AfjbqOegIkx17VjVj_0ThjyPp; FPGSID=1.1723123453.1723123753.G-B8E0BLEGRH.-otcjU7vaLtEZpVT0pODNA; _ga_PT43NBSMZN=GS1.3.1723123434.1.1.1723123761.48.0.0; _clsk=v6bos0%7C1723123762037%7C3%7C0%7Cz.clarity.ms%2Fcollect; _ga_B8E0BLEGRH=GS1.1.1723123453.1.1.1723123762.0.0.1012662251; _ga=GA1.1.671066005.1723123434; FCNEC=%5B%5B%22AKsRol9Jq4S3GBzylTrgfG2DXG2yARw4s0BKsfcN7K_t8DLSM_IbzdOfP5Z3B53oDYh8jLtneDvv5kSnBixplVcf7cop1lDwZKFevMuD91ygIMh7iajxQ75KeZVXOBjUQqUKjro3Fe2QxSjC1DjAE9zlA4Hfex0SVw%3D%3D%22%5D%5D'

}
    
data = requests.get(url,headers=header)

data.encoding = "utf-8"

data=data.text

soup = BeautifulSoup(data,'html.parser')

allnews = soup.find(class_='news_list')

news = allnews.find(class_='list')

items = news.find_all('li')

for row in items:
    print(row)
    break