from Singleton import Singleton
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

class CreateBrowser(metaclass = Singleton):

    def __init__(self):
        self.chrome_options = Options()
        self.chrome_options.add_argument('start-maximized')
        self.chrome_options.add_argument('--disable-notifications')
        
    def createBrowser(self):
        return webdriver.Chrome(options= self.chrome_options)

class Browser:

    def __init__(self):
        self.browser = CreateBrowser().createBrowser()

    def goPage(self, pageURL: str):
        self.browser.get(pageURL)

    def quitBrowser(self):
        self.browser.quit()

    def getElement(self, elementName: str):
        return self.browser.find_element_by_xpath(elementName)

    def getElements(self, elementName: str):
        return self.browser.find_elements_by_xpath(elementName)

    def clickElement(self, elementName: str):
        button = self.getElement(elementName)
        self.browser.execute_script("arguments[0].click();", button)

    def clickElementWithoutXpath(self, element):
        self.browser.execute_script("arguments[0].click();", element)

    def readFile(fileName: str):
        inf_file = open(fileName, 'r',encoding='utf-8')
        lines = inf_file.readlines()
        inf = [line.rstrip('\r\n') for line in lines]
        inf_file.close()
        return inf

    def waitElement(self, elementName: str, waitTime: int):
        WebDriverWait(self.browser, waitTime).until(EC.presence_of_element_located((By.XPATH, elementName)))

