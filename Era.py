from RealEstate import RealEstate
from ExcelConnection import ExcelConnection
from selenium.common.exceptions import NoSuchElementException
from Browser import Browser
from time import sleep

class Era(RealEstate):
    def __init__(self) -> None:
        self.offices = []
        self.agents = []
        self.browser = Browser()

    def getOffices(self) -> None:
        self.browser.goPage("https://www.eragayrimenkul.com.tr/tr-TR/Offices/Search?sorting=monthlytransactionturnover%2c2")
        sleep(3)
        endPageButton = self.browser.getElements("//a[@class='page-link']")[-1]
        self.browser.clickElementWithoutXpath(endPageButton)
        sleep(2)
        pageCount = int(self.browser.getElements("//a[@class='page-link']")[-2].text)
        eraUrl = "https://www.eragayrimenkul.com.tr/tr-TR/Offices/Search?sorting=monthlytransactionturnover%2c2&pager_p="
        
        for page in range(1, pageCount + 1):
            self.browser.goPage(eraUrl + str(page))
            sleep(2)
            officeInfo = self.browser.getElements("//h3[@class='title']/a")

            [self.offices.append(office.get_attribute("href")) for office in officeInfo]

        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Ofis Adresi", "Il - Ilce", "Telefon 1", "Telefon 2", "E-mail", "Website", "Facebook", "Twitter", "Linkedin", "Youtube"))

        for office in self.offices:
            self.browser.goPage(office)
            sleep(2)
            officeName = self.browser.getElement("//h3[@class='text-center text-md-left']").text
            officeAddress = self.browser.getElement("//div[@class=' d-flex justify-content-between bg-secondary text-white p-4']/div[@class='content']/p").text
            officeProvince = self.browser.getElement("//div[@class='col-12 col-md-3 text-white text-center text-lg-right']/h3").text
            infos = self.browser.getElements("//ul[@class='buttons']/li/div/a")

            infosLength = len(infos)
            
            if infosLength == 3:
                email = infos[0].text
                phoneNumber = infos[1].text
                wpNumber = infos[2].text
            
            if infosLength == 2:
                if "+" in infos[0].text:
                    phoneNumber = infos[0].text
                    wpNumber = "-"
                    email = "-"

                if "@" in infos[0].text:
                    email = infos[0].text
                    phoneNumber = "-"
                    wpNumber = "-"

                if "+" in infos[1].text:
                    wpNumber = infos[1].text

                if "@" in infos[1].text:
                    email = infos[1].text

            if infosLength == 1:
                if "+" in infos[0].text:
                    phoneNumber = infos[0].text
                    wpNumber = "-"
                    email = "-"
                else:
                    email = infos[0].text
                    phoneNumber = "-"
                    wpNumber = "-"

            socialMedia = self.browser.getElements("//div[@class='d-flex socials mt-2']/a")
            facebook, twitter, linkedin, youtube = socialMedia[0].get_attribute("href"), socialMedia[1].get_attribute("href"), socialMedia[2].get_attribute("href"), socialMedia[3].get_attribute("href")
            print(officeName, officeAddress, officeProvince, phoneNumber, wpNumber, email, office, facebook, twitter, linkedin, youtube, sep = "|")
            sheet.append((officeName, officeAddress, officeProvince, phoneNumber, wpNumber, email, office, facebook, twitter, linkedin, youtube))
            book.save("./files/era_offices.xlsx")
        book.close()

    def getAgents(self) -> None:
        self.browser.goPage("https://www.eragayrimenkul.com.tr/tr-TR/OfficeUsers/Search?sorting=monthlytransactionturnover%2c2&pager_p=1")
        sleep(3)
        self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
        endPageButton = self.browser.getElement("//li[@class='page-item']")
        endPageButton.click()
        sleep(3)
        pageCount = int(self.browser.getElements("//ul[@class='pagination']/li")[-2].text)
        eraUrl = "https://www.eragayrimenkul.com.tr/tr-TR/OfficeUsers/Search?sorting=monthlytransactionturnover%2c2&pager_p="
        
        for page in range(1, pageCount + 1):
            self.browser.goPage(eraUrl + str(page))
            sleep(3)
            agentInfo = self.browser.getElements("//div[@class='content']/a")
            [self.agents.append(agent.get_attribute("href")) for agent in agentInfo]
        
        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Rolu", "Adi Soyadi", "Telefon 1", "Telefon 2", "E-mail", "Website", "Facebook", "Instagram", "Linkedin", "Twitter", "Youtube"))

        for agent in self.agents:
            self.browser.goPage(agent)
            sleep(2)
            officeName = self.browser.getElement("//small[@class='form-text  mb-2']/a").text.split("@")[1]
            degree = self.browser.getElement("//small[@class='form-text  mb-2']").text.split(" @")[0] 
            name = self.browser.getElement("//h3[@class='text-center text-md-left']").text
            infos = self.browser.getElements("//ul[@class='buttons']/li/div/a")
            
            infosLength = len(infos)
            
            if infosLength == 3:
                email = infos[0].text
                phoneNumber = infos[1].text
                wpNumber = infos[2].text
            
            if infosLength == 2:
                if "+" in infos[0].text:
                    phoneNumber = infos[0].text
                    wpNumber = "-"
                    email = "-"

                if "@" in infos[0].text:
                    email = infos[0].text
                    phoneNumber = "-"
                    wpNumber = "-"

                if "+" in infos[1].text:
                    wpNumber = infos[1].text

                if "@" in infos[1].text:
                    email = infos[1].text

            if infosLength == 1:
                if "+" in infos[0].text:
                    phoneNumber = infos[0].text
                    wpNumber = "-"
                    email = "-"
                else:
                    email = infos[0].text
                    phoneNumber = "-"
                    wpNumber = "-"

            socialMedia = self.browser.getElements("//div[@class='d-flex socials mt-2']/a")
            facebook, instagram, linkedin, twitter, youtube = socialMedia[0].get_attribute("href"), socialMedia[1].get_attribute("href"), socialMedia[2].get_attribute("href"), socialMedia[3].get_attribute("href"), socialMedia[4].get_attribute("href")
            print(officeName, degree, name, phoneNumber, wpNumber, email, agent, facebook, instagram, linkedin, twitter, youtube, sep = "|")
            sheet.append((officeName, degree, name, phoneNumber, wpNumber, email, agent, facebook, instagram, linkedin, twitter, youtube))
            book.save("./files/era_agents.xlsx")
        book.close()