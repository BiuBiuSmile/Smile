# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

from bs4 import BeautifulSoup

import requests

url ="https://www.setn.com/ViewAll.aspx"

header = {
    
'user-agent':
'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36',
'cookie':
'__htid=4d38d2f7-1d67-4564-9729-72708c04c2d7; _ht_em=1; AviviD_uuid=707aaa74-f5a3-463e-a1ce-c69609a3daa8; webuserid=8b059091-0e0a-0379-b8a5-0e60704af111; userKey=86f37e9f-04dd-47b2-a5be-0260b5a2c2ff; AviviD_already_exist=0; AviviD_sw_version=1.0.868.210701; _ga=GA1.1.566075274.1723118381; ch_tracking_uuid=1; _ht_47b240=1; _fbp=fb.1.1723118381640.738130595639724552; AviviD_tid_rmed=1; _cc_id=c3c2ec5bfcc9e854b615c5545a14d2cd; panoramaId_expiry=1723723181835; panoramaId=8eb9e09c0d518a35b2044861ec3c16d539381239d2d67f260882c8bddafb3a71; panoramaIdType=panoIndiv; stid=566075274.1723118381; stid2=566075274.1723118381; __gads=ID=539d9c39c9fbc8b0:T=1723118381:RT=1723118381:S=ALNI_MbAmTxBp4L0NT6V44-CxriBTSwuCA; __gpi=UID=00000eb935809f0a:T=1723118381:RT=1723118381:S=ALNI_MaBOlpAj5WjTanjLIHYGIxwO90XNQ; __eoi=ID=23604bdfa695161c:T=1723118381:RT=1723118381:S=AA-AfjYWexOnNxf2BvL__42897Tz; Privacy=true; _ht_hi=1; AviviD_refresh_uuid_status=2; show_avivid_native_subscribe=2; _ttd_sync=1; oid=%257B%2522oid%2522%253A%2522dbdd7660-5574-11ef-bdb2-0242ac130002%2522%252C%2522ts%2522%253A-62135596800%252C%2522v%2522%253A%252220201117%2522%257D; dable_uid=36724559.1723114585440; _ss_pp_id=349145a0eb94a1b74761723085784940; _im_vid=01J4RXSX0DSCR9SSQDY4AQ48BV; vuukle_geo_region={%22country_code%22:%22TW%22%2C%22os%22:%22Windows%22%2C%22device%22:%22Tablet%22%2C%22browser%22:%22Chrome%22}; uid-s=ccdb157-a695-43a2-ba7b-f9d65f12c723; vsid=0ec3cb1d-be89-48e6-9b93-49130401691f; jiyakeji_uuid=dba48540-557d-11ef-839b-31ec17522d9f; truvid_protected={"val":"c","level":1,"geo":"TW","timestamp":1723118461}; cto_bundle=xn3iuV9qYXJTNjVpbmdLd1E0SWZIamFJMVJqaUwlMkZqMW5SV0JXUW93aG1PSFFqdSUyRmoyeEk5UGl1WWR3WW9OVXFKcmJjdmVqY3pQVWs3ejMzQSUyQlRwcGpnUHhZWTdDM3ZEN3ZGY2o5UFYlMkI0YlI1SFJHeW9CcFJCc0NLeHJVYkhpY24lMkJEaiUyRlZXaFpQTWZVbzkxTVVGRzVUUEhOYlElM0QlM0Q; cto_bidid=iJZPEl9DYzF0TXdUa2JCTVE4WXhBZlFMREZBQ3pPSXRVSlltZUU2bnRzUzFkNFM5MExjMGJuOGEzWEd0c00wOGJZZmNNQnRyUm5TamNYQmtSdlFwVEpTbHVoQm9TMjNiMm1kNDhNMiUyQnlNaWhzOGowJTNE; _td=6d638b1f-6310-454a-8db0-0fcffc84184d; _ga_ZEP9LRMW9Y=GS1.1.1723118459.1.1.1723118490.29.0.0; _ga_8NJ3QZRCY6=GS1.1.1723118381.1.1.1723118524.59.0.0; _ga_54ZJR9ZRH0=GS1.1.1723118381.1.1.1723118524.59.0.0; FCNEC=%5B%5B%22AKsRol9jrCkE6UKcIf_kNnE1WReJ6NmOrvN4BD1Yp2zaLuAu-vPsXPfXa-JoM7xxU8iUBr7gt_9WJqV8HCtKBCmVrfklRXYV24ruRBRgKrRYRcniRoW-MJWnnoU91WhQu-u2Su-ZCileiMi_1n-6D98l-GUexf_Nnw%3D%3D%22%5D%5D'
     }
    
data = requests.get(url,headers=header)

data.encoding = "utf-8"

data=data.text

soup = BeautifulSoup(data,'html.parser')

allnews=soup.find(id='NewsList')

times = allnews.find_all('time')
h3s = allnews.find_all('h3')

for i in range(len(times)):
    
    time = times[i].text
    title=h3s[i].text
    link = h3s[i].find('a').get('href')
    
    if not('https' in link):
        link='www.setn.com'+link
    print('標題:',title)
    print('時間:',time)
    print(link)
    print(title)
    print()
    
   
     
    


