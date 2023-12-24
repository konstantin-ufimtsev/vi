from openpyxl import workbook
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import logging
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

logging.basicConfig(
    filename='log.log',
    filemode='w',
    level=logging.INFO,
    encoding='utf-8',
    datefmt="[%d-%m-%Y %H:%M:%S]",
    format='%(asctime)s %(levelname)-8s [%(filename)s:%(lineno)d %(message)s]'
)

class vi_parsing():
    def __init__(self, url:str, pages:int):
        self.url = url
        self.pages=pages

    def write_to_file(self, data:list):
        filename = "result.xlsx"
        try:
            wb = load_workbook(filename)
            page = wb.active
            for row in data:
                page.append(row)
            wb.save(filename)
            logging.info(f'Данные записаны в файл: количество строк - {len(data)}')
            print(f'Данные записаны в файл: количество строк - {len(data)}')
        except Exception as ex:
            logging.info(f"Ошибка записи в фай: {ex}")
            print(f"Ошибка записи в фай: {ex}")
            
    def get_url(self):
        try:
            self.driver.get(self.url)
            logging.info(f'В драйвер загружена страница {self.url}')
            time.sleep(5)
            WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "[class='button-icon -no-text']")))
            print(f'В драйвер загружена страница {self.url}')
        except Exception as ex:
            logging.info(f'Ошибка загрузки страницы {ex}')
            print('Станица загружена в драйвер')
                

    def setup(self):
        options = webdriver.ChromeOptions()
        options.add_argument("user-agent=Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:84.0) Gecko/20100101 Firefox/84.0")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument('--ignore-certificate-errors')
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)
        #options.add_argument("--headless=new")
        options.add_argument("--disable-logging")
        self.driver = webdriver.Chrome(options=options)
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        'source': '''
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
            delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
        '''
        })
        self.driver.maximize_window()
        
    """
    def pagination(self):
        count = 0
        
        while self.driver.find_elements(By.CSS_SELECTOR, "[xmlns='http://www.w3.org/2000/svg']") and count < self.pages:
            count +=1
            time.sleep(5)
            WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "[xmlns='http://www.w3.org/2000/svg']")))
            self.parse_page()
            self.driver.find_element(By.CSS_SELECTOR, "[xmlns='http://www.w3.org/2000/svg']").click()
            logging.info(f"Отработала пагинация номер {count}")
            print(f"Отработала пагинация номер {count}")
            time.sleep(5)
    """
    def parse_page(self):
        items = self.driver.find_elements(By.CSS_SELECTOR, "[data-qa='products-tile']")
        logging.info(f"Найдено количество товаров: {len(items)}")
        print(f"Найдено количество товаров: {len(items)}")
        page_data = []
        for item in items:
            temp_list = []
            try:
                name = item.find_element(By.CSS_SELECTOR, "[class='typography text v2 -no-margin']").text
            except Exception as ex:
                print("Ошибка считывания названия")
            try:
                article = item.find_element(By.CSS_SELECTOR, "[data-qa='product-code-text']").text.split()[1]
            except Exception as ex:
                print("Ошибка считывания артикула")
            try:
                befor_price = float(item.find_element(By.CSS_SELECTOR, "[data-qa='product-price-old-value']").text.replace(" ", "").replace("р.", ""))
                befor_price = round(befor_price, 2)
            except Exception as ex:
                befor_price = ""
            try:
                after_price = float(item.find_element(By.CSS_SELECTOR, "[data-qa='product-price-current']").text.replace(" ", "").replace("р.", ""))
                after_price = round(after_price, 2)
            except Exception as ex:
                after_price = ""
            temp_list = [article, name, befor_price, after_price]
            print(temp_list)
            page_data.append(temp_list)
        self.write_to_file(page_data)
    
    def parse(self):
        self.setup()
        self.get_url()
        self.parse_page()
        


if __name__ == "__main__":
    
    pages = 10
    
    urls1 = [f"https://www.vseinstrumenti.ru/category/kotly-otopleniya-2780/page{i}/" for i in range(1, 10)]
    urls2 = [f"https://www.vseinstrumenti.ru/category/akkumulyatornyj-instrument-2392/page{i}" for i in range(1, 10)]
    
    urls = urls1 + urls2
    
    start_time = time.perf_counter()
    
    for url in urls:
        vi_parsing(url, pages=10).parse()
    end_time = time.perf_counter()
    print('Время выполнения парсинга Леруа мерлен:', round((float(end_time - start_time) / 60), 1), ' минут!')