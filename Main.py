from RealEstateFactory import RealEstateFactory
from PyQt5 import QtWidgets
from UserInterface import Ui_MainWindow
import qdarkstyle

class UI(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.buttonCheck)
        self.ui.checkBox_19.toggled.connect(self.isCheckedAll)
        self.checkBoxes = [
                            self.ui.checkBox, self.ui.checkBox_2, self.ui.checkBox_3, self.ui.checkBox_4,
                            self.ui.checkBox_5, self.ui.checkBox_6, self.ui.checkBox_7, self.ui.checkBox_8,
                            self.ui.checkBox_9, self.ui.checkBox_10, self.ui.checkBox_11, self.ui.checkBox_12,
                            self.ui.checkBox_13, self.ui.checkBox_14, self.ui.checkBox_15, self.ui.checkBox_16,
                            self.ui.checkBox_17, self.ui.checkBox_18 
                          ]

    def runFunctions(self):
        realEstateObject = RealEstateFactory()

        if self.ui.checkBox.isChecked() == True:
            realEstateObject.createRealEstateObject("kwturkiye").getOffices()

        if self.ui.checkBox_2.isChecked() == True:
            realEstateObject.createRealEstateObject("kwturkiye").getAgents()

        if self.ui.checkBox_3.isChecked() == True:
            realEstateObject.createRealEstateObject("era").getOffices()

        if self.ui.checkBox_4.isChecked() == True:
            realEstateObject.createRealEstateObject("era").getAgents()

        if self.ui.checkBox_5.isChecked() == True:
            realEstateObject.createRealEstateObject("cb").getOffices()

        if self.ui.checkBox_6.isChecked() == True:
            realEstateObject.createRealEstateObject("cb").getAgents()

        if self.ui.checkBox_7.isChecked() == True:
            realEstateObject.createRealEstateObject("premar").getOffices()

        if self.ui.checkBox_8.isChecked() == True:
            realEstateObject.createRealEstateObject("premar").getAgents()

        if self.ui.checkBox_9.isChecked() == True:
            realEstateObject.createRealEstateObject("kristalturkiye").getOffices()

        if self.ui.checkBox_10.isChecked() == True:
            realEstateObject.createRealEstateObject("kristalturkiye").getAgents()

        if self.ui.checkBox_11.isChecked() == True:
            realEstateObject.createRealEstateObject("turyap").getOffices()

        if self.ui.checkBox_12.isChecked() == True:
            realEstateObject.createRealEstateObject("turyap").getAgents()

        if self.ui.checkBox_13.isChecked() == True:
            realEstateObject.createRealEstateObject("tamnokta").getOffices()

        if self.ui.checkBox_14.isChecked() == True:
            realEstateObject.createRealEstateObject("tamnokta").getAgents()

        if self.ui.checkBox_15.isChecked() == True:
            realEstateObject.createRealEstateObject("realtyworld").getOffices()
        
        if self.ui.checkBox_16.isChecked() == True:
            realEstateObject.createRealEstateObject("realtyworld").getAgents()
        
        if self.ui.checkBox_17.isChecked() == True:
            realEstateObject.createRealEstateObject("centuryglobal").getOffices()
        
        if self.ui.checkBox_18.isChecked() == True:
            realEstateObject.createRealEstateObject("centuryglobal").getAgents()
        
        self.close()

    def isCheckedAll(self):
        [checkBox.setChecked(self.ui.checkBox_19.isChecked()) for checkBox in self.checkBoxes]

    def showMessage(self, message, title):
        mes = QtWidgets.QMessageBox()
        mes.setIcon(QtWidgets.QMessageBox.Information)
        mes.setText(message)
        mes.setInformativeText("")
        mes.setWindowTitle(title)
        mes.setStandardButtons(QtWidgets.QMessageBox.Ok)
        mes.exec_()
    
    def buttonCheck(self):
        controls = [checkBox.isChecked() for checkBox in self.checkBoxes]
        if controls.count(False) == 18:
            self.showMessage("Lütfen seçim yaptığınızdan emin olunuz..", "Seçim Yapılmadı!") 
        else:
            self.runFunctions()

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    window = UI()
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    window.show()
    app.exec_()