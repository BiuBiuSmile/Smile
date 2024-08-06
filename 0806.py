# -*- coding: utf-8 -*-
"""
Created on Tue Aug  6 19:02:25 2024

@author: USER
"""

import requests 
import io
import xml.sax
class BusHandler(xml.sax.ContentHandler):
    def startElement(self, tag, attr):
      if tag == "Route":
        print(attr["nameZh"])
        print(attr['ddesc'])
        print(attr['departureZh'])
        print()
if __name__ == '__main__':
    
    parser=xml.sax.make_parser()
    bus = BusHandler()
    parser.setContentHandler(bus)
    url="https://ibus.tbkc.gov.tw/xmlbus/StaticData/GetRoute.xml"
    data= requests.get(url)
    data.encoding="utf-8"
    data= data.text
    f=io.StringIO(data)
    parser.parse(f)       