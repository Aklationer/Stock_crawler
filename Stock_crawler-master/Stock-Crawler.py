import requests
import threading
from lxml import etree
import openpyxl
from openpyxl.styles import Font
class Stock:
    def __init__(self,*numbers):
        self.numbers=numbers
        self.data=[]

    def thread(self):
        for number in self.numbers:
            response=threading.Thread(target=Stock.crawler,args=(self,'https://tw.stock.yahoo.com/quote/'+number))
            response.start()
    def crawler(self,url):
            temp=[]
            response=requests.get(url)
            html=etree.HTML(response.text)
            temp.append(html.xpath('//time/span[2]/text()')[0])
            temp.append(html.xpath('//div[@class="D(f) Ai(c) Mb(6px)"]/h1/text()')[0])
            prices=html.xpath('//ul[@class="D(f) Fld(c) Flw(w) H(192px) Mx(-16px)"]/li')
            for price in prices:
                temp.append(price.xpath('./span[2]/text()')[0])
            self.data.append(temp)
            if len(self.data) == len(self.numbers):
                self.save(self.data)

    def save(self,stocks):
        wb=openpyxl.Workbook()
        sheet=wb.create_sheet('Stock',0)
        sheet.append(('時間','股票代號','成交','開盤','最高','最低','均價','成交金額(億)','昨收','漲跌幅','漲跌','總量','昨量','振幅'))
        for stock in stocks:
            sheet.append(stock)
        wb.save("stock_data.xlsx")
  
s =Stock('2451','2454','2369','3189','3034','2342','2303','2302','2451','3545','2340')
s.thread()