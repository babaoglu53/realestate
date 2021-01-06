from selenium.common.exceptions import NoSuchElementException
from ExcelConnection import ExcelConnection
from RealEstate import RealEstate
from Browser import Browser
from time import sleep

class TamNokta(RealEstate):

    def __init__(self) -> None:
        self.offices = []
        self.agents = []
        self.browser = Browser()

    def getOffices(self) -> None:
        self.browser.goPage("https://www.tamnokta.com.tr/temsilciliklerimiz/")
        sleep(2)
        
        self.offices = [office.get_attribute("href") for office in self.browser.getElements("//a[@class='agent-list-link']")]

        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Ofis Adresi", "Telefon 1", "Telefon 2", "E-mail", "Website", "Sosyal Medya"))
        
        for office in self.offices:
            self.browser.goPage(office)
            sleep(2)
            officeName = self.browser.getElement("//div[@class='agent-profile-header']/h1").text
            officeAddress = self.browser.getElement("//div[@class='agent-map']/address").text
            phoneNumbers = self.browser.getElements("//ul[@class='list-unstyled']/li/a/span")
            phoneNumber_1 = phoneNumbers[0].text
            try:
                phoneNumber_2 = phoneNumbers[1].text
            except IndexError:
                phoneNumber_2 = "-"
            try:
                email = self.browser.getElement("//ul[@class='list-unstyled']/li[@class='email']/a").text
            except NoSuchElementException:
                email = "-"
            websiteLink = self.browser.getElements("//ul[@class='list-unstyled']/li/a")[-1].get_attribute("href")
            socialMedias = ", ".join([socialMedia.get_attribute("href") for socialMedia in self.browser.getElements("//div[@class='agent-social-media']/span/a")])
            profiles_t = self.browser.getElements("//div[@class='d-flex xxs-column']/h2/a")
            for profile_t in profiles_t:
                self.agents.append(profile_t.get_attribute("href"))
            print(officeName, officeAddress, phoneNumber_1, phoneNumber_2, email, websiteLink, socialMedias, sep = "|")
            sheet.append((officeName, officeAddress, phoneNumber_1, phoneNumber_2, email, websiteLink, socialMedias))
            book.save("./files/tamNokta_offices.xlsx")
        book.close()

    def getAgents(self) -> None:
        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Rolu", "Adi Soyadi", "Telefon 1", "Telefon 2", "E-mail", "Website", "Sosyal Medya"))
        for agent in self.agents:
            self.browser.goPage(agent)
            sleep(2)
            officeName = self.browser.getElement("//p[@class='agent-list-position']/a").text
            degree = self.browser.getElement("//p[@class='agent-list-position']").text.split(" en")[0]
            name = self.browser.getElement("//div[@class='agent-profile-header']/h1").text
            phoneNumbers = self.browser.getElements("//ul[@class='list-unstyled']/li/a/span")
            phoneNumber_1 = phoneNumbers[0].text
            try:
                phoneNumber_2 = phoneNumbers[1].text
            except IndexError:
                phoneNumber_2 = "-"
            try:
                email = self.browser.getElement("//ul[@class='list-unstyled']/li[@class='email']/a").text
            except NoSuchElementException:
                email = "-"
            websiteLink = self.browser.getElements("//ul[@class='list-unstyled']/li/a")[-1].get_attribute("href")
            socialMedias = ", ".join([socialMedia.get_attribute("href") for socialMedia in self.browser.getElements("//div[@class='agent-social-media']/span/a")])
            if socialMedias == "": socialMedias = "-"
            print(officeName, degree, name, phoneNumber_1, phoneNumber_2, email, websiteLink, socialMedias, sep = "|")
            sheet.append((officeName, degree, name, phoneNumber_1, phoneNumber_2, email, websiteLink, socialMedias))
            book.save("./files/tamNokta_agents.xlsx")
        print("Process completed..")
        book.close()