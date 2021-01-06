from RealEstate import RealEstate
from ExcelConnection import ExcelConnection
from Browser import Browser
from time import sleep

class KristalTurkiye(RealEstate):
    def __init__(self) -> None:
        self.offices = []
        self.agents = []
        self.provinces = []
        self.browser = Browser()

    def getOffices(self) -> None:
        self.kristalTurkiyeGetProvinces()
        self.browser.goPage("https://kristalturkiye.com/ofislerimiz/")
        sleep(2)
        offices_t = self.browser.getElements("//div[@class='pic']/a")
        
        for office_t in offices_t:
            self.offices.append(office_t.get_attribute("href"))
        
        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Ofis Adresi", "Il - Ilce", "Telefon Numarasi", "E-mail", "Website"))

        for office, province in zip(self.offices, self.provinces):
            self.browser.goPage(office)
            sleep(2)
            self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
            officeName = self.browser.getElement("//h1[@data-aos='zoom-in']").text
            infos = self.browser.getElements("//table[@class='table bordernone']/tbody/tr")
            officeAddress = infos[0].text
            if officeAddress == "": officeAddress = "-"
            officeProvince = province
            if (officeProvince == "") or (officeProvince == "/") : officeProvince = "-"
            phoneNumber = infos[1].text
            email = self.browser.getElements("//table[@class='table bordernone']/tbody/tr/td")[-1].text
            if email == "": email = "-"
            websiteLink = office
            agents_t = self.browser.getElements("//div[@class='pics']/a")
            for agent_t in agents_t:
                self.agents.append(agent_t.get_attribute("href"))
            print(officeName, officeAddress, officeProvince, phoneNumber, email, websiteLink, sep = "|")
            sheet.append((officeName, officeAddress, officeProvince, phoneNumber, email, websiteLink))
            book.save("./files/kristalTurkiye_offices.xlsx")
        book.close()

    def getAgents(self) -> None:
        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Rolu", "Adi Soyadi", "Telefon Numarasi", "E-mail", "Website"))
        for agent in self.agents:
            self.browser.goPage(agent)
            self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
            self.browser.waitElement("//h1[@data-aos='zoom-in']", 5)
            officeName = self.browser.getElement("//h1[@data-aos='zoom-in']").text
            if officeName == "":
                sleep(2)
                self.browser.execute_script("window.scrollTo(document.body.scrollHeight, 0);")
                officeName = self.browser.getElement("//h1[@data-aos='zoom-in']").text
            degree = self.browser.getElement("//h3[@class='text-center title']/small").text
            name = self.browser.getElement("//h3[@class='text-center title']").text.split("\n")[0]
            phoneNumber = self.browser.getElements("//p[@class='text-center']/a/b")[0].text
            email = self.browser.getElements("//p[@class='text-center']/a/b")[1].text
            websiteLink = agent
            print(officeName, degree, name, phoneNumber, email, websiteLink, sep = "|")
            sheet.append((officeName, degree, name, phoneNumber, email, websiteLink))
            book.save("./files/kristalTurkiye_agents.xlsx")
        book.close()

    def kristalTurkiyeGetProvinces(self):
        self.browser.goPage("https://kristalturkiye.com/ofislerimiz/")
        sleep(2)
        self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
        sleep(2)   
        provinces_t = self.browser.getElements("//div[@class='team-content']//span[@class='post']")
        for province_t in provinces_t:
            self.provinces.append(province_t.text)
