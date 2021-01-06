from RealEstate import RealEstate
from ExcelConnection import ExcelConnection
from Browser import Browser
from time import sleep

class Turyap(RealEstate):
    def __init__(self) -> None:
        self.offices = []
        self.agents = []
        self.agentNames = []
        self.agentPositions = []
        self.agentNumbers_1 = []
        self.agentNumbers_2 = []
        self.browser = Browser()

    def getOffices(self) -> None:
        self.browser.goPage("http://www.turyap.com.tr/ofislerimiz")
        sleep(3)
        totalNumber = int(self.browser.getElement("//div[@id='top_pane']//span[@id='total_number']").text)
        currentNumber = 0
        while totalNumber != currentNumber:
            sleep(2)
            offices_t = self.browser.getElements("//div[@class='pull-right']/a")
            [self.offices.append(office_t.get_attribute("href")) for office_t in offices_t]
            self.browser.waitElement("//div[@id='top_pane']//i[@class='icon icon-round icon-next']",5)
            self.browser.clickElement("//div[@id='top_pane']//i[@class='icon icon-round icon-next']")
            currentNumber = int(self.browser.getElement("//div[@id='top_pane']//span[@id='current_number']").text.split("- ")[1])
        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Ofis Adi", "Ofis Adresi", "Il", "Ilce", "Telefon 1", "Telefon 2", "Fax", "Website"))

        for office in self.offices:
            self.browser.goPage(office)
            sleep(2)
            officeName = self.browser.getElement("//div[@class='panel-body']/div[@class='title']").text
            addressInfo = self.browser.getElements("//div[@class='address']/div[@class='line']")
            officeAddress, officeProvince, officeDistrict = addressInfo[0].text, addressInfo[1].text, addressInfo[2].text
            numbers = self.browser.getElements("//div[@class='data-wrapper']//span[@dir='ltr']")
            phoneNumber1, phoneNumber2 = numbers[0].text, numbers[1].text
            try:
                faxNumber = numbers[2].text
            except:
                faxNumber = "-"
                pass
            websiteLink = self.browser.getElement("//span[@class='pull-left']/a").get_attribute("href")

            agentNames_t = self.browser.getElements("//div[@class='data-wrapper col-xs-12 col-sm-12 col-md-12 col-lg-12']/div[@class='name text-center']")
            [self.agentNames.append(agentName_t.text) for agentName_t in agentNames_t]
            
            agentPositions_t  = self.browser.getElements("//div[@class='data-wrapper col-xs-12 col-sm-12 col-md-12 col-lg-12']/div[@class='position']")
            [self.agentPositions.append(agentPosition_t.text) for agentPosition_t in agentPositions_t]
            
            agentNumbers_1_t  = self.browser.getElements("//div[@class='data-wrapper col-xs-12 col-sm-12 col-md-12 col-lg-12']/div[@class='mobile-phone']/span")
            [self.agentNumbers.append(agentNumber_1_t.text) for agentNumber_1_t in agentNumbers_1_t]
            
            agentNumbers_2_t  = self.browser.getElements("//div[@class='data-wrapper col-xs-12 col-sm-12 col-md-12 col-lg-12']/div[@class='phone']/span")
            [self.agentNumbers.append(agentNumber_2_t.text) for agentNumber_2_t in agentNumbers_2_t]

            print(officeName, officeAddress, officeDistrict, officeProvince, phoneNumber1, phoneNumber2, faxNumber, websiteLink, sep = "|")
            sheet.append((officeName, officeAddress, officeDistrict, officeProvince, phoneNumber1, phoneNumber2, faxNumber, websiteLink))
            book.save("./files/turyap_offices.xlsx")
        book.close()

    def getAgents(self) -> None:
        book = ExcelConnection().connectToExcel()
        sheet = book.active
        sheet.append(("Rolu", "Adi Soyadi", "Telefon 1", "Telefon 2"))
        for agentPosition, agentName, agentNumber_1, agentNumber_2 in zip(self.agentNames, self.agentPositions, self.agentNumbers_1, self.agentNumbers_2):
            print(agentPosition, agentName, agentNumber_1, agentNumber_2, sep = "|")
            sheet.append((agentPosition, agentName, agentNumber_1, agentNumber_2))
            book.save("./files/turyap_agents.xlsx")
        book.close()