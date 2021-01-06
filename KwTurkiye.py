from RealEstate import RealEstate
from ExcelConnection import ExcelConnection
from selenium.common.exceptions import NoSuchElementException
from Browser import Browser
from time import sleep

class KwTurkiye(RealEstate):
    def __init__(self) -> None:
        self.offices = []
        self.agents = []
        self.browser = Browser()

    def getOffices(self) -> None:
        self.browser.goPage("https://www.kwturkiye.com/officeagentsearch/")
        sleep(3)
        pageCount = int(self.browser.getElements("//a[@class='ng-binding']")[-2].text)
        for page in range(pageCount):
            sleep(3)
            officeInfo = self.browser.getElements("//a[@class='office-name ng-binding']")
            for office in officeInfo:
                self.offices.append(office.get_attribute("href"))
            try:
                self.browser.clickElement("//li[@class='pagination-next ng-scope']/a")
            except NoSuchElementException:
                print("Ofisler cekildi..")

        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Ofis Adresi", "Il- Ilce", "Telefon Numarasi", "Website", "Yonetim"))
        for office in self.offices:
            try:
                self.browser.goPage(office)
                self.browser.waitElement("//div[@ng-if='vm.officeStaff.length > 0']//div[@class='agent-languages']/div/div", 5)
                officeName = self.browser.getElement("//div[@class='location-office']//div[@class='col-xs-12']/h4/span").text
                officeAddress_t = self.browser.getElement("//div[@class='location-office']//div[@class='address ng-scope']/div").text
                officeAddress = " ".join(officeAddress_t.split("\n"))
                officeProvince = officeAddress_t.split("\n")[-1]
                
                try:
                    officeNumber = self.browser.getElement("//div[@class='officeagent-comm-item ng-scope']/a[@class='icon-btn-phone']").get_attribute("alt")
                except:
                    officeNumber = "-"
                
                try:
                    websiteLink = self.browser.getElement("//div[@ng-if='vm.office.personalWebsite']//a").get_attribute("href")
                except:
                    websiteLink = "-"

                government_t = self.browser.getElements("//div[@ng-if='vm.officeStaff.length > 0']//div[@class='agent-languages']/div/div")
                government_t = [i.text for i in government_t]
                government = ", ".join(government_t).replace("\n"," ")

                print(officeName, officeAddress, officeProvince, officeNumber, websiteLink, government, sep = "|")
                sheet.append((officeName, officeAddress, officeProvince, officeNumber, websiteLink, government))
                book.save("./files/kwTurkiye_offices.xlsx")
                
            except:
                print("Ofis gecildi..")

        book.close()

    def getAgents(self) -> None:
        self.browser.goPage("https://www.kwturkiye.com/agents")
        sleep(3)
        pageCount = int(self.browser.getElements("//a[@class='ng-binding']")[-2].text) ##1 eksigi kere basacak sonraki butonuna
        for page in range(pageCount):
            sleep(3)
            agentInfo = self.browser.getElements("//a[@class='agent-name ng-binding']")
            for agent in agentInfo:
                self.agents.append(agent.get_attribute("href"))
            try:
                self.browser.clickElement("//li[@class='pagination-next ng-scope']/a")
            except NoSuchElementException:
                print("Kisiler cekildi..")
        counter = 1
        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Adi Soyadi", "Ofis Adi", "Ofis Adresi", "Telefon 1", "Telefon 2", "Website"))

        for agent in self.agents:
            try:
                self.browser.goPage(agent)
                self.browser.waitElement("//div[@class='col-xs-12 form-group']/div[@class='photo-agent']/h2/a", 3)
                name = self.browser.getElements("//div[@class='col-xs-12 form-group']/div[@class='photo-agent']/h2/a")[1].text
                officeName = self.browser.getElement("//div[@class='office-details']/h3/a").text
                officeAddress = self.browser.getElement("//div[@class='office-details']/div[@class='form-group officeagent-addr']/span").text
                try:
                    wpNumber = self.browser.getElements("//a[@class='icon-btn-whatsapp']")[1].get_attribute("href").split("phone=")[1]
                except:
                    wpNumber = "-"
                    pass
                try:
                    phoneNumber = self.browser.getElements("//a[@class='icon-btn-phone']")[2].get_attribute('href').split("tel:")[1]
                except:
                    phoneNumber = "-"
                    pass

                print(str(counter), name, officeName, officeAddress, wpNumber, phoneNumber, agent, sep = "|")
                sheet.append((name, officeName, officeAddress, wpNumber, phoneNumber, agent))
                book.save("./files/kwTurkiye_agents.xlsx")
                counter += 1
            except:
                print("Profil gecildi..")          
        book.close()