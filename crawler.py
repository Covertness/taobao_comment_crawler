import sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from pyquery import PyQuery as pq
import openpyxl
from time import sleep

class tmall_infos:

    #对象初始化
    def __init__(self):
        url = 'https://login.taobao.com/'
        self.url = url

        options = webdriver.ChromeOptions()
        options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
        options.add_experimental_option('excludeSwitches', ['enable-automation'])

        self.browser = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
        self.wait = WebDriverWait(self.browser, 60)
    
    def login(self):
        self.browser.get(self.url)

        self.browser.implicitly_wait(30)
        taobao_name = self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.site-nav-bd > ul.site-nav-bd-l > li#J_SiteNavLogin > div.site-nav-menu-hd > div.site-nav-user > a.site-nav-login-info-nick ')))

    def detail(self, url):
        self.browser.get(url)

        self.browser.implicitly_wait(30)
        self.browser.find_element("xpath", '//*[@id="root"]/div/div[2]/div[1]/div[2]/div/div[2]/div/div[2]/div[1]/div/div[4]/div').click()
        self.browser.implicitly_wait(30)

        # scr1 = self.browser.find_element("xpath", '/html/body/div[7]/div/div[2]/div/div[3]')

        # self.swipe_down(20, scr1)

        input("把评论滑到最后点击回车...")

        html = self.browser.page_source

        doc = pq(html)
        comment_items = doc('.Comment--root--3_CJ07v').items()

        comment_data = []
        for item in comment_items:
            meta = item.find('.Comment--meta--1AM9IDf').text().replace('\n',"").replace('\r',"")
            comment = item.find('.Comment--content--22pGCmW').text().replace('\n',"").replace('\r',"")
            print("---------------------------------------------------------------------------------------")
            print("meta:", meta)
            print("comment:", comment)
            print("\n")

            meta_array = meta.split('·')
            meta_array.append(comment)

            comment_data.append(meta_array)
        
        return comment_data

    def swipe_down(self, second, scrollDiv):
        self.browser.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight", scrollDiv)
        sleep(second)

    def gen_xlsx(self, data):
        wb = openpyxl.Workbook()

        sheet = wb.active
  
        for row in data: 
            sheet.append(row)
            
        wb.save(f'tmall_comments_{datetime.now().strftime("%Y_%m_%d_%H_%M_%S")}.xlsx')

if __name__ == "__main__":
    if len(sys.argv) <= 1:
        print("请输入 url")
        sys.exit(-1)
    
    url = sys.argv[1]

    a = tmall_infos()
    a.login()

    data = a.detail(url)
    a.gen_xlsx(data)
