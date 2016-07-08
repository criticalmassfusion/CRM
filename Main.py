from __future__ import print_function

import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from sub.ui_contacts import Ui_contacts
from sub.ui_commentaire import Ui_Commentaire
from sub.ui_evenement import Ui_Rappel
from sub.ui_info import Ui_Info
from sub.ui_bad_info import Ui_BadInfo
from sub.ui_area_memo import Ui_areaMemo
from sub.ui_attention import Ui_Attention
from sub.google import Google

from docxtpl import *
import os
from openpyxl import Workbook
from openpyxl import load_workbook
import re
import datetime
from dateutil import parser
import webbrowser
import subprocess

import httplib2
from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


# Constantes et Variables Globales
# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES =            'https://www.googleapis.com/auth/calendar'
APPLICATION_NAME =  'Google Calendar API'

# Chemin de tous les fichiers utilisés
FICHIER_EXCEL =         'save/Contacts.xlsx'
CLIENT_SECRET_FILE =    'save/client_secret.json'
TEMPLATE_FICHE_VISITE = 'temp/my_word_template.docx'
TEMPLATE_DEVIS_SIMPLE = 'temp/DEVIS 1sur1.docx'
TEMPLATE_DEVIS_DOUBLE = 'temp/DEVIS 1sur1.docx'


# Sous-Classes de GUI
class Attention(QtWidgets.QDialog):
    def __init__(self):
        super(Attention, self).__init__()

        # Set up the user interface from Designer.
        self.ui = Ui_Attention()
        self.ui.setupUi(self)

        self.ui.addButton.clicked.connect(self.ok)
        self.ui.cancelButton.clicked.connect(self.refus)
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnTopHint)

    def ok(self):
        pass

    def refus(self):
        pass


# Classe Principale
class Main(Ui_contacts):
    def __init__(self):
        Ui_contacts.__init__(self)
        self.setupUi(main)
        self.main = main
        self.main.setWindowState(QtCore.Qt.WindowMaximized)
        # self.main.setWindowFlags(QtCore.Qt.WindowMaximizeButtonHint | QtCore.Qt.WindowMinimizeButtonHint)

        # Initialisations
        self.google = Google()

    def init_recherche(self):
        # Initialisation champ tableau et sélecteur
        self.searchSelector.addItems(self.excel.champs)
        self.resultTable.setColumnCount(len(self.excel.champs))
        self.resultTable.setHorizontalHeaderLabels(self.excel.champs)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    main = QtWidgets.QWidget()
    ui = Main()
    main.show()
    sys.exit(app.exec_())