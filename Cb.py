from RealEstate import RealEstate
from ExcelConnection import ExcelConnection
from Browser import Browser
from time import sleep

class Cb(RealEstate):
    def __init__(self) -> None:
        self.offices = []
        self.agents = []
        self.browser = Browser()

    def getOffices(self) -> None:
        self.browser.goPage("https://www.cb.com.tr/tr-TR/Offices/Search?sorting=name,1")
        sleep(3)
        endPageButton = self.browser.getElements("//ul[@class='pagination']/li")[-1]
        self.browser.clickElementWithoutXpath(endPageButton)
        sleep(2)
        pageCount = int(self.browser.getElements("//ul[@class='pagination']/li")[-2].text)
        cbUrl = "https://www.cb.com.tr/tr-TR/Offices/Search?sorting=name%2c1&pager_p="
        
        for page in range(1, pageCount+1):
            self.browser.goPage(cbUrl + str(page))
            sleep(2)
            officeInfo = self.browser.getElements("//h3[@class='title']/a")
            [self.offices.append(office.get_attribute("href")) for office in officeInfo]
        
        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Ofis Adresi", "Il- Ilce", "E-mail", "Telefon 1", "Telefon 2", "Website", "Facebook", "Twitter", "Linkedin", "Youtube"))

        for office in self.offices:
            self.browser.goPage(office)
            sleep(2)
            officeName = self.browser.getElement("//h3[@class='text-center text-md-left']").text
            officeAddress = self.browser.getElement("//div[@class=' d-flex justify-content-between bg-secondary text-white p-4']/div[@class='content']/p").text
            officeProvince = self.browser.getElement("//div[@class='col-12 col-md-3 text-white text-center text-lg-right']/h3").text
            infos = self.browser.getElements("//ul[@class='buttons']/li/div/a")
            try: 
                email = infos[0].text
            except:
                email = "-"
                
            try:
                phoneNumber = infos[1].text
            except:
                phoneNumber = "-"
                
            try:
                wpNumber = infos[2].text
            except:
                wpNumber = "-"
                            
            
            socialMedia = self.browser.getElements("//div[@class='d-flex socials mt-2']/a")
            facebook, twitter, linkedin, youtube = socialMedia[0].get_attribute("href"), socialMedia[1].get_attribute("href"), socialMedia[2].get_attribute("href"), socialMedia[3].get_attribute("href")
            print(officeName, officeAddress, officeProvince, email, phoneNumber, wpNumber, office, facebook, twitter, linkedin, youtube, sep = "|")
            sheet.append((officeName, officeAddress, officeProvince, email, phoneNumber, wpNumber, office, facebook, twitter, linkedin, youtube))
            book.save("./files/cb_offices.xlsx")
        book.close()
    
    def getAgents(self) -> None:
        self.browser.goPage("https://www.cb.com.tr/tr-TR/OfficeUsers/Search?sorting=monthlytransactionturnover%2c2&pager_p=1")
        sleep(3)
        self.browser.execute_script("window.scrollTo(0, document.body.scrollHeight);var lenOfPage=document.body.scrollHeight;return lenOfPage;")
        endPageButton = self.browser.getElement("//li[@class='page-item']")
        endPageButton.click()
        sleep(3)
        pageCount = int(self.browser.getElements("//ul[@class='pagination']/li")[-2].text)
        print(pageCount)
        cbUrl = "https://www.cb.com.tr/tr-TR/OfficeUsers/Search?sorting=monthlytransactionturnover%2c2&pager_p="
        
        for page in range(1, pageCount + 1):
            self.browser.goPage(cbUrl + str(page))
            sleep(3)
            agentInfo = self.browser.getElements("//div[@class='content']/a")
            [self.agents.append(agent.get_attribute("href")) for agent in agentInfo]

        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Adi Soyadi", "Unvan", "Telefon 1", "Telefon 2", "Website", "E-mail", "Facebook","Instagram", "Linkedin", "Twitter", "Youtube"))

        for agent in self.agents:
            self.browser.goPage(agent)
            sleep(2)
            officeName = self.browser.getElement("//small[@class='form-text  mb-2']/a").text.split("@")[1]
            name = self.browser.getElement("//h3[@class='text-center text-md-left']").text
            degree = self.browser.getElement("//small[@class='form-text  mb-2']").text.split(" @")[0]
            infos = self.browser.getElements("//ul[@class='buttons']/li/div/a")
            try: 
                email = infos[0].text
            except:
                email = "-"
                
            try:
                phoneNumber = infos[1].text
            except:
                phoneNumber = "-"
                
            try:
                wpNumber = infos[2].text
            except:
                wpNumber = "-"

            socialMedia = self.browser.getElements("//div[@class='d-flex socials mt-2']/a")
            facebook, instagram, linkedin, twitter, youtube = socialMedia[0].get_attribute("href"), socialMedia[1].get_attribute("href"), socialMedia[2].get_attribute("href"), socialMedia[3].get_attribute("href"), socialMedia[4].get_attribute("href")
            print(officeName, name, degree, phoneNumber, wpNumber, agent, email, facebook, instagram, linkedin, twitter, youtube, sep = "|")
            sheet.append((officeName, name, degree, phoneNumber, wpNumber, agent, email, facebook, instagram, linkedin, twitter, youtube))
            book.save("./files/cb_agents.xlsx")
        book.close()