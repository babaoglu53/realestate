from RealEstate import RealEstate
from ExcelConnection import ExcelConnection
from selenium.common.exceptions import NoSuchElementException
from Browser import Browser
from time import sleep

class Premar(RealEstate):
    def __init__(self) -> None:
        self.offices = []
        self.agents = []
        self.provinces = []
        self.provinces_a = []
        self.browser = Browser()

    def getOffices(self) -> None:
        premarUrl = "https://www.premar.com.tr/ofisler"
        self.browser.goPage(premarUrl)
        pageCount = int(self.browser.getElements("//ul[@class='pagination']/li")[-2].text)
        
        for page in range(1, pageCount + 1):
            self.browser.goPage(premarUrl + "?page=" + str(page))
            sleep(2)
            offices_t = self.browser.getElements("//div[@class='content p-2']/a")
            provinces_t = self.browser.getElements("//span[@class='city']")
            
            for office_t, province_t in zip(offices_t, provinces_t):
                self.offices.append(office_t.get_attribute("href"))
                self.provinces.append(province_t.text)
        
        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Ofis Adresi", "Il - Ilce", "Telefon 1", "Telefon 2", "E-mail", "Website"))
        
        for office, province in zip(self.offices, self.provinces):
            try:
                self.browser.goPage(office)
                sleep(2)
                infos = self.browser.getElements("//span[@class='content d-block text-muted']")
                officeName = self.browser.getElement("//h1[@class='name']").text
                officeAddress = infos[0].text
                officeProvince = province
                phoneNumber = infos[2].text
                phoneNumber2 = infos[3].text
                if phoneNumber2 == "": phoneNumber2 = "-"
                email = infos[4].text
                if "@" not in email: email = "-"
                websiteLink = infos[1].text

                print(officeName, officeAddress, officeProvince, phoneNumber, phoneNumber2, email, websiteLink, sep = "|")
                sheet.append((officeName, officeAddress, officeProvince, phoneNumber, phoneNumber2, email, websiteLink))
                book.save("./files/premar_offices.xlsx")
            except NoSuchElementException:
                print("Bu ofis gecildi..")
        book.close()

    def getAgents(self) -> None:
        premarUrl = "https://www.premar.com.tr/danismanlar"
        self.browser.goPage(premarUrl)
        sleep(2)
        pageCount = int(self.browser.getElements("//ul[@class='pagination']/li")[-2].text)

        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Rolu", "Adi Soyadi", "Telefon 1", "Telefon 2", "E-mail"))

        for page in range(1, pageCount + 1):
            self.browser.goPage(premarUrl + "?page=" + str(page))
            sleep(2)
            agents_t = self.browser.getElements("//div[@class='content p-2']/a")
            provinces_t = self.browser.getElements("//span[@class='city']")

            for agent_t, province_t in zip(agents_t, provinces_t):
                self.agents.append(agent_t.get_attribute("href"))
                self.provinces_a.append(province_t.text)

        for agent, province in zip(self.agents, self.provinces_a):
            self.browser.goPage(agent)
            self.browser.waitElement("//a[@class='phone']", 5)
            officeName = self.browser.getElements("//ol[@class='breadcrumb']/li")[1].text
            degree = self.browser.getElement("//h6[@class='user-title']").text
            agentName = self.browser.getElement("//h5[@class='user-name mt-2']").text
            infos = self.browser.getElements("//a[@class='phone contact-link']")
            phoneNumber = infos[0].text
            phoneNumber2 = self.browser.getElement("//a[@class='phone']").text
            try:
                email = infos[1].text
            except IndexError:
                email = "-"
            print(officeName, degree, agentName, phoneNumber, phoneNumber2, email, sep = "|")
            sheet.append((officeName, degree, agentName, phoneNumber, phoneNumber2, email))
            book.save("./files/premar_agents.xlsx")
        book.close()