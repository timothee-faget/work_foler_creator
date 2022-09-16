# Imports

import xlsxwriter
import sys
import os
import json

from PyQt5 import QtWidgets
from datetime import date

from window.wind_Main import Ui_FolderCreator

# Code

def initialisation():
    config = json.load(open('config.json'))
    global GT_folder, GT_subfolders, Romax_folder, folders_types, default_subfolders, separator_index
    GT_folder = config['GT_folder']
    Romax_folder = config['ROMAX_folder']
    separator_index = config['separator_index']
    GT_subfolders = []
    for sub in config['GT_subfolders']:
        GT_subfolders.append(sub['subfolder'])
    folders_types = []
    for sub in config['folder_types']:
        folders_types.append(sub['type'])
    default_subfolders = []
    for sub in config['default_subfolders']:
        default_subfolders.append(sub['subfolder'])


class MainWindow(QtWidgets.QMainWindow, Ui_FolderCreator):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        self.setupUi(self)

        self.pushButton_creer.setEnabled(False)

        self.comboBox_type.addItems(folders_types)
        self.comboBox_type.insertSeparator(1)
        self.comboBox_type.insertSeparator(separator_index)

        self.conditions = [0, 0]

        self.pushButton_creer.clicked.connect(self.clickedCreer)
        self.pushButton_ann.clicked.connect(self.clickedAnn)
        self.comboBox_type.currentTextChanged.connect(self.projectTypeChanged)
        self.lineEdit_nom.textEdited.connect(self.projectNameChanged)

    def clickedAnn(self):
        self.close()

    def clickedCreer(self):
        folder_adress, folder_num = self.getNewFolderNumber()
        folder_to_create = f"{folder_adress}\{folder_num}_{self.lineEdit_nom.text()}"
        os.mkdir(folder_to_create)
        for sub in default_subfolders:
            sub_folder_to_create = f"{folder_adress}\{folder_num}_{self.lineEdit_nom.text()}\{sub}"
            os.mkdir(sub_folder_to_create)
        self.createDescriptionFile(folder_to_create)
        if self.checkBox_excel.isChecked():
            self.createExcelSheets(folder_to_create)
        self.close()

    def projectTypeChanged(self):
        if self.comboBox_type.currentText() != "-":
            self.conditions[0] = 1
        else:
            self.conditions[0] = 0
        self.checkConditions()

    def projectNameChanged(self):
        if self.lineEdit_nom.text() != "":
            self.conditions[1] = 1
        else:
            self.conditions[1] = 0
        self.checkConditions()

    def checkConditions(self):
        if sum(self.conditions) == 2:
            self.pushButton_creer.setEnabled(True)
        else:
            self.pushButton_creer.setEnabled(False)

    def getNewFolderNumber(self):
        if self.comboBox_type.currentIndex() < separator_index:
            full_folder_adress = GT_folder + GT_subfolders[self.comboBox_type.currentIndex()-2]
        elif self.comboBox_type.currentIndex() == separator_index + 1:
            full_folder_adress = Romax_folder
        folders = []
        for file in os.listdir(full_folder_adress):
            d = os.path.join(full_folder_adress, file)
            if os.path.isdir(d):
                folders.append(file)
        if not folders:
            new_num  = 1
        else:
            nums = []
            for f in folders:
                nums.append(int(f.split("_")[0]))
            new_num = max(nums) + 1
        if new_num < 10:
            new_num = f"0{new_num}"
        else:
            new_num = str(new_num)
        
        return full_folder_adress, new_num

    def createDescriptionFile(self, folder):
        today = date.today()
        d1 = today.strftime("%d/%m/%Y")
        with open(rf"{folder}\00_Lisez-moi.txt", 'w') as f:
            f.write("--Projet:\n")
            f.write(f"{self.lineEdit_nom.text()}\n\n")
            f.write("--Créé le:\n")
            f.write(f"{today.strftime('%d/%m/%Y')}\n\n")
            f.write("--Description projet:\n")
            f.write(f"{self.textEdit_description.toPlainText()}")

    def createExcelSheets(self, folder):
        workbook = xlsxwriter.Workbook(rf"{folder}\02_Data\01_Data_Modele.xlsx")
        workbook.close()
        workbook = xlsxwriter.Workbook(rf"{folder}\04_Analyse_Resultats\01_Analyse.xlsx")
        workbook.close()
        workbook = xlsxwriter.Workbook(rf"{folder}\04_Analyse_Resultats\02_Resultats.xlsx")
        workbook.close()

def main():
    initialisation()
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    win = MainWindow()
    win.show()
    sys.exit(app.exec())

# Launcher

if __name__=="__main__":
    main()