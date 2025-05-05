from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import xml.etree.ElementTree as et
import bs4 as bs
import pandas as pd
from datetime import datetime as dt
from pathlib import Path
import json
import os

class marks:
    def __init__(self):
        secrets_path = Path(__file__).parent / "secrets.json"

        with open(secrets_path, "r") as f:
            secrets = json.load(f)
        
        hall_tickets_env =os.getenv("HALL_TICKETS","[]")
        self.lis1 =json.loads(hall_tickets_env)
        self.results_url =os.getenv("RESULTS_URL")
        print(self.results_url,self.lis1)
        self.lis2=[]
        self.initialize_browser()
        

    def initialize_browser(self):
        options=webdriver.ChromeOptions()
        options.add_argument("--headless=new")  # New headless mode (Chrome 109+)
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")


        self.browser = webdriver.Chrome(options=options)

        self.browser.get(self.results_url)
        self.browser.implicitly_wait(10)
        self.req()


    def req(self):
        for ele in self.lis1:
            self.browser.find_element(By.NAME,"txtHallTicketNo").send_keys(ele)
            self.browser.find_element(By.NAME,"btnSubmit").click()
            sleep(2)
            pagesource=self.browser.page_source
            self.parse(pagesource)
            sleep(2)
            self.browser.find_element(By.NAME,"txtHallTicketNo").clear()
            sleep(2)

        self.to_excel()
    
    def parse(self,source):
        lis1=[]
        file=bs.BeautifulSoup(source,"lxml")
        rows=file.find_all('table',class_="style1")
        tds=rows[0].find_all('td')
        for ele in tds:
            lis1.append(ele.text.strip("\t").strip(" ").strip("\n").strip(" ").strip("\n"))

        lis1=lis1[:-1]
        fir=lis1[-1]
        sec=lis1[-2]
        lis1=lis1[:-2]
        fir=fir.split(":")
        
        for ele in range(len(fir)):
            if ele==0:
                lis1.append(fir[ele].strip("\n").strip(" "))
            else:
                lis1.append(fir[ele].strip(" \n"))
        sec=sec.split(":")
        for ele in range(len(sec)):
            if ele==0:
                lis1.append(sec[ele].strip("\n").strip(" "))
            else:
                lis1.append(sec[ele].strip(" \n"))
        print(lis1)
        self.parse2(lis1)

    def parse2(self,lis1):
        i=j=0
        self.dict1={}
        lis2=[]
        lis3=[]
        for ele in range(len(lis1)):
            lis1[ele]="".join(lis1[ele].split("*")[0].strip())
            if ele<=7:
                if ele%2==0:
                    key=lis1[ele]
                else:
                    val=lis1[ele]
                    self.dict1[key]=val
            elif 12<=ele<30:
                if '\xa0' in lis1[ele]:
                    lis1[ele]=lis1[ele].replace('\xa0',"")
                lis2.append(lis1[ele])
            elif ele>=30:
                lis3.append(lis1[ele])
            
        while(i<len(lis2)):
            v=lis2[i:i+3]
            self.dict1[f'{v[0]}_Marks']=int(v[1])
            self.dict1[f'{v[0]}_Grade']=v[2]
            i+=3
        self.dict1[lis3[0]]=lis3[1]
        self.dict1[lis3[2]]=int(lis3[3])
        self.lis2.append(self.dict1)
    
    def to_excel(self):
        df=pd.DataFrame(data=self.lis2)
        timestamp=dt.now().strftime("%Y%m%d_%H%M%S")
        filename= f"marks_{timestamp}.xlsx"
        df.to_excel(filename,index=False)
                        
       
if __name__ =="__main__":
    marks()


