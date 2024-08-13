# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

from bs4 import BeautifulSoup

import requests

url ="https://news.cts.com.tw/real/index.html"

header = {
    
'user-agent':
'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36',
'cookie':
'_gid=GA1.3.59876476.1723114573; _fbp=fb.2.1723114573408.123534135966611075; AviviD_uuid=707aaa74-f5a3-463e-a1ce-c69609a3daa8; webuserid=5d44730c-e59d-a0e2-e345-12b38e6fac30; ch_tracking_uuid=1; _ht_47b240=1; __htid=4d38d2f7-1d67-4564-9729-72708c04c2d7; _ht_em=1; AWSALB=Xn9+c5cc6Z+T5uEDib0oL//n8eCGlDmbTJA/tR7pb7lWT+8o1CcT4P47TQkvPIoQLu+ZzV3Z6MnUUB9VsGzOwT7WBvrPkvoFtPojZU7jNTbxRzRbCiSqgbTJl9Xf; AWSALBCORS=Xn9+c5cc6Z+T5uEDib0oL//n8eCGlDmbTJA/tR7pb7lWT+8o1CcT4P47TQkvPIoQLu+ZzV3Z6MnUUB9VsGzOwT7WBvrPkvoFtPojZU7jNTbxRzRbCiSqgbTJl9Xf; AviviD_refresh_uuid_status=2; ISMD5VERSION=1; CFFPCKUUID=6555-jVWhQIHP9nZ0pAUxWqHS8w8bpriUyQ5H; CFFPCKUUIDMAIN=6106-VVHTD1YCTG3vLVQGgwfccEvcU1Y3etoa; FPUUID=6106-d1c47c251736e3392bfb3a1e3981df33; _ht_hi=1; _ht_50ef57=1; _ss_pp_id=349145a0eb94a1b74761723085784940; _td=aa3a3635-ee57-450c-9121-67da6bf4fb25; dable_uid=36724559.1723114585440; __gads=ID=5261fbeefbc467e1:T=1723114581:RT=1723115032:S=ALNI_MauV0FBqMseCHNyubFbSzW4-XVmyw; __gpi=UID=00000eb92d822fd3:T=1723114581:RT=1723115032:S=ALNI_MYjGp1GWJP4eNVlLH3yJ-ynIkp3Dg; __eoi=ID=43ac4395391fa141:T=1723114581:RT=1723115032:S=AA-AfjahDc2DwYRifOsOM0vRRkx8; _ga=GA1.3.2096810252.1723114573; _ga_FWQF4JLMB1=GS1.3.1723114575.1.1.1723115108.53.0.0; _ga_F6LZM62QYC=GS1.3.1723114575.1.1.1723115108.53.0.0; _ga_B5S0TX9D32=GS1.1.1723114573.1.1.1723115110.50.0.1608176925; FCNEC=%5B%5B%22AKsRol_eic4EFupkU4GvJswVnr3hnpJLSr-gaG2RZopB7RdmTqUxtQCQ56im33-vlt2keMsBaLBX1dEmqFx_NN5r5lgcIKw_nK5RlP32junnNavcUVdHBFS4BExIhyaHl1E5TKCmUHdXSnruPtDJzagzwv2fkk61sQ%3D%3D%22%5D%5D'
     }
    
data = requests.get(url,headers=header)

data.encoding = "utf-8"

data=data.text

soup = BeautifulSoup(data,'html.parser')

allnews=soup.find(id='newslist-top')

news = allnews.find_all('a')

for row in news:
    link = row.get('href')
    title = row.get('title')
    h3=row.find('h3')
    if h3 != None:
        h3=h3.text
    else:
        h3=""
        
    img=row.find('img')
    
    if img !=None:
          photo = img.get('data-src')
    else:
          photo = ""
    print('連結:',link) 
    print('標題:',title)
    print(h3)
    print('圖片:',photo)
     
    


