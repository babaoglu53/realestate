from RealEstate import RealEstate
from ExcelConnection import ExcelConnection
from selenium.common.exceptions import NoSuchElementException
from Browser import Browser
from time import sleep

class RealtyWorld(RealEstate):
    def __init__(self) -> None:
        self.officeNames = []
        self.officeAddress = []
        self.officeProvinces = []
        self.phoneNumbers = []
        self.websiteLinks = []
        self.emails = []
        self.agents = []
        self.browser = Browser()

    def getOffices(self) -> None:
        self.browser.goPage("https://www.realtyworld.com.tr/tr/ofisler")
        sleep(2)
        self.browser.clickElement("//div[@class='formBlock select']//button[@type='submit']")
        sleep(5)
        self.officeNames = self.browser.getElements("//div[@class='col-sm-6 col-md-4']/header/h3/a")
        self.officeAddress_a = self.browser.getElements("//div[@class='col-12 col-md-5']/p")
        self.officeProvinces = self.browser.getElements("//div[@class='col-sm-6 col-md-4']/header/h4")
        self.phoneNumbers = self.browser.getElements("//div[@class='col-sm-6 col-md-4']/header/h5")
        self.websiteLinks = self.browser.getElements("//div[@class='col-12 col-md-5']/p/a")
        self.emails = self.browser.getElements("//div[@class='col-sm-6 col-md-4']/header/h5/a")

        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Ofis Adresi", "Il - Ilce", "Numara", "Website", "E-mail"))

        for officeName, officeAddress, officeProvince, phoneNumber, websiteLink, email in zip(self.officeNames, self.officeAddress_a, self.officeProvinces, self.phoneNumbers, self.websiteLinks, self.emails):
            print(officeName.text.replace("\n", " "), officeAddress.text.split("\n")[0], officeProvince.text, phoneNumber.text.split("\n")[0], websiteLink.get_attribute("href"), email.get_attribute("href").split("mailto:")[1], sep = "|")
            phoneNumber = phoneNumber.text.split("\n")[0]
            if phoneNumber[0] != "+":
                phoneNumber = "-"
            sheet.append((officeName.text.replace("\n", " "), officeAddress.text.split("\n")[0], officeProvince.text, phoneNumber, websiteLink.get_attribute("href"), email.get_attribute("href").split("mailto:")[1]))
            book.save("./files/realtyWorld_offices.xlsx")
        book.close()

        
    def getAgents(self) -> None:
        self.browser.goPage("https://www.realtyworld.com.tr/tr/ofisler")
        sleep(2)
        self.browser.clickElement("//a[@href='#tab2']")
        sleep(1)
        
        searchButton = self.browser.getElements("//div[@class='formBlock select']//button[@type='submit']")[1]
        self.browser.clickElementWithoutXpath(searchButton)
        self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
        sleep(2)
        self.agents = self.browser.getElements("//div[@class='col-sm-6 col-md-4']/header/h3/a") #.get_attribute("href")
        self.agents = [agent.get_attribute("href") for agent in self.agents]

        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Rolu", "Adi Soyadi", "Telefon 1", "Telefon 2", "E-mail"))
        
        for agent in self.agents:
            try:
                self.browser.goPage(agent)
                self.browser.waitElement("//div[@class='col-sm-8 col-md-9 col-lg-10']/h3", 5)
                officeName = self.browser.getElement("//div[@class='col-sm-8 col-md-9 col-lg-10']/p/a").text
                agentDegree = self.browser.getElement("//div[@class='col-sm-8 col-md-9 col-lg-10']/p/b").text
                agentName = self.browser.getElement("//div[@class='col-sm-8 col-md-9 col-lg-10']/h3").text
                agentInfos = self.browser.getElements("//div[@class='col-sm-8 col-md-9 col-lg-10']/p")[1].text.split("\n")
                phoneNumber1 = agentInfos[0]
                if "+" not in phoneNumber1: phoneNumber1 = agentInfos[1]
                phoneNumber2 = agentInfos[2]
                if "+" not in phoneNumber2: phoneNumber2 = "-"
                email = self.browser.getElements("//div[@class='col-sm-8 col-md-9 col-lg-10']/p/a")[1].text
                print(officeName, agentDegree, agentName, phoneNumber1, phoneNumber2, email, sep = "|")
                sheet.append((officeName, agentDegree, agentName, phoneNumber1, phoneNumber2, email))
                book.save("./files/realtyWorld_agents.xlsx")
            except:
                print("Bu ofis gecildi..")
                pass
        book.close()