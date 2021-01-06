from RealEstate import RealEstate
from ExcelConnection import ExcelConnection
from Browser import Browser
from time import sleep


class CenturyGlobal(RealEstate):
    
    def __init__(self) -> None:
        self.offices = []
        self.agents = []
        self.browser = Browser()
    
    def getOffices(self) -> None:
        self.browser.goPage("https://www.century21global.com/tr/emlak-ofisleri/Turkey")
        sleep(2)
        print("Ofisler cekiliyor..")
        self.offices = [office.get_attribute("href") for office in self.browser.getElements("//a[@class='search-result-photo']")]
        print("Ofisler cekildi..")
        
        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Ofis Adresi", "Il - Ilce", "Telefon"))

        for office in self.offices:
            self.browser.goPage(office)
            officeName = self.browser.getElement("//h2[@class='color-gold-highlight']").text
            infos = self.browser.getElement("//div[@class='detail-data-primary']").text
            officeAddress = " ".join(infos.split("\n")[:-1])
            if "Telefon" in officeAddress: officeAddress = officeAddress.split(" Telefon numarası:")[0]
            province = infos.split("\n")[1]
            try:
                phoneNumber = infos.split("Telefon numarası: ")[1]
                if "F" in phoneNumber: phoneNumber = phoneNumber.split("\n")[0]
            except IndexError:
                phoneNumber = "-"
            print(officeName, officeAddress, province, phoneNumber, sep = "|")
            sheet.append((officeName, officeAddress, province, phoneNumber))
            book.save("./files/centuryGlobal_offices.xlsx")
            agentsButton = self.browser.getElements("//button[@class='btn btn-primary']")[0]
            self.browser.clickElementWithoutXpath(agentsButton)
            sleep(2)
            agents_t = self.browser.getElements("//a[@class='search-result-info']")
            for agent_t in agents_t:
                self.agents.append(agent_t.get_attribute("href"))
        book.close()

    def getAgents(self) -> None:
        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Adi Soyadi", "Telefon 1", "Telefon 2"))
        for agent in self.agents:
            self.browser.goPage(agent)
            officeName = self.browser.getElement("//h2[@class='color-gold-highlight']/a").text
            name = self.browser.getElements("//h2[@class='color-gold-highlight']")[0].text
            phoneNumbers = self.browser.getElement("//div[@class='detail-data-primary']").text
            phoneNumber_1 = phoneNumbers.split("\n")[-2].split("Cep telefonu: ")[1]
            phoneNumber_2 = phoneNumbers.split("\n")[-1].split("Telefon numarası: ")[1]
            print(officeName, name, phoneNumber_1, phoneNumber_2, sep = "|")
            sheet.append((officeName, name, phoneNumber_1, phoneNumber_2))
            book.save("./files/centuryGlobal_agents.xlsx")
        print("Kisiler tamamlandi..\nIslem tamamlandi..")
        book.close()
