from PyQt6 import QtCore, QtGui, QtWidgets
from PyQt6.QtGui import QIcon
from PyQt6.QtGui import QFont
import os
from os.path import exists
import win32com.client
import webbrowser

#if u want to contribute to my project
linkgit = 'https://github.com/parreira7/photoshoPY'
def link():
        webbrowser.open_new(linkgit)
#{} []
#funçoes
def abrirpsd():
        psApp = win32com.client.Dispatch("Photoshop.Application")
        psApp.Open(r"C:\Users\Administrator\Desktop\Python\photoshoPY\thumnail_P.psd")
        doc = psApp.Application.ActiveDocument

def saveaspng():
        save = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb') #selecionando a funçao de exportar 
        save.Format = 13 # formato que no caso, o png vale 13 na tabela
        save.PNG8 = False # seta como png 24-bit
        pngfile = (r"C:\Users\Administrator\Desktop\Python\photoshoPY\thumb.png")
        psApp = win32com.client.Dispatch("Photoshop.Application")
        doc = psApp.Application.ActiveDocument 
        doc.Export(ExportIn=pngfile, ExportAs=2, Options=save) #chamando a funçao com tudo já setado (path, formato etc)
def sair():
        arquivo_existe = os.path.exists(r"C:\Users\Administrator\Desktop\Python\photoshoPY\thumb.png") #mudar no seu pc
        print(arquivo_existe)#printa se existe ou nao
        if (arquivo_existe == True): #se existir vai fechar
                psApp = win32com.client.Dispatch("Photoshop.Application")
                doc = psApp.Application.ActiveDocument
                doc = psApp.Application.ActiveDocument.Close(1) #salvar e fechar o .psd
                psApp.Quit() #fechar o photoshop 


        
class Ui_Window(object):
        def setupUi(self, Window):
        #janela
                Window.setObjectName("Window")
                Window.resize(650, 300)
                Window.setMinimumSize(QtCore.QSize(650, 300))
                Window.setMaximumSize(QtCore.QSize(650, 300))
                icon = QtGui.QIcon()
                icon.addPixmap(QtGui.QPixmap("ps.png"), QtGui.QIcon.Mode.Normal, QtGui.QIcon.State.On)
                Window.setWindowIcon(icon)
                Window.setStyleSheet("background-color:#474044")
                
                self.centralwidget = QtWidgets.QWidget(Window)
                self.centralwidget.setAutoFillBackground(False)
                self.centralwidget.setObjectName("centralwidget")
                
                self.abrirps = QtWidgets.QPushButton(self.centralwidget)
                self.abrirps.setGeometry(QtCore.QRect(20, 90, 51, 61))
                self.abrirps.setStyleSheet("background-color:#FFBF46; color:#474044;     border-width: 2px;\n"
        "    border-radius: 10px")
                self.abrirps.setText("")
                self.abrirps.setObjectName("abrirps")
                self.abrirps.setIcon(QIcon('ps.png'))
                self.abrirps.setIconSize(QtCore.QSize(50, 50))
                self.abrirps.clicked.connect(abrirpsd)
                
                self.caixatexto = QtWidgets.QTextEdit(self.centralwidget)
                self.caixatexto.setGeometry(QtCore.QRect(90, 20, 351, 141))
                self.caixatexto.viewport().setProperty("cursor", QtGui.QCursor(QtCore.Qt.CursorShape.IBeamCursor))
                self.caixatexto.setMouseTracking(False)
                self.caixatexto.setStyleSheet("background-color:#FFBF46; font: 16pt \"Largo\"; color:#474044;     border-width: 2px;\n"
        "    border-radius: 10px;")
                self.caixatexto.setObjectName("caixatexto")
                
                self.sair = QtWidgets.QPushButton(self.centralwidget)
                self.sair.setGeometry(QtCore.QRect(10, 20, 71, 41))
                font = QtGui.QFont()
                font.setFamily("Largo")
                font.setPointSize(12)
                font.setBold(False)
                font.setItalic(False)
                font.setWeight(50)
                self.sair.setFont(font)
                self.sair.setStyleSheet("font: 12pt \"Largo\";  background-color:#ED1C24; color:#fff;     border-width: 2px;\n"
        "    border-radius: 10px;\n"
        "")
                self.sair.clicked.connect(sair)
                self.sair.setObjectName("sair")
                
                self.salvarpng = QtWidgets.QPushButton(self.centralwidget)
                self.salvarpng.setGeometry(QtCore.QRect(90, 170, 161, 41))
                font = QtGui.QFont()
                font.setFamily("Largo")
                font.setPointSize(12)
                font.setBold(False)
                font.setItalic(False)
                font.setWeight(50)
                self.salvarpng.clicked.connect(saveaspng)
                self.salvarpng.setFont(font)
                self.salvarpng.setAutoFillBackground(False)
                self.salvarpng.setStyleSheet("font: 12pt \"Largo\";  background-color:#2BA84A; color:#fff;     border-width: 2px;\n"
        "    border-radius: 10px;")
                self.salvarpng.setObjectName("salvarpng")
                
                self.inserir = QtWidgets.QPushButton(self.centralwidget)
                self.inserir.setGeometry(QtCore.QRect(260, 170, 81, 41))
                font = QtGui.QFont()
                font.setFamily("Largo")
                font.setPointSize(12)
                font.setBold(False)
                font.setItalic(False)
                font.setWeight(50)
                self.inserir.setFont(font)
                self.inserir.setStyleSheet("font: 12pt \"Largo\";  background-color:#FFBF46; color:#474044;     border-width: 2px;\n"
        "    border-radius: 10px;\n"
        "")
                self.inserir.setObjectName("inserir")
                
                def nome1():
                        psApp = win32com.client.Dispatch("Photoshop.Application")
                        doc = psApp.Application.ActiveDocument
                        layerText = doc.ArtLayers["NOME1"] #selecionando a layer com o nome especifico
                        texto_layer = layerText.TextItem #listar como texto
                        texto_layer.contents = [caixatexto]
                self.nome1 = QtWidgets.QPushButton(self.centralwidget)
                self.nome1.clicked.connect(nome1)
                self.nome1.setGeometry(QtCore.QRect(460, 10, 161, 51))
                self.nome1.setStyleSheet("font: 16pt \"Largo\"; background-color:#FFBF46; color:#474044;  border-width: 2px;\n"
        " border-radius: 10px;")
                self.nome1.setObjectName("nome1")
                
                self.nome2 = QtWidgets.QPushButton(self.centralwidget)
                self.nome2.setGeometry(QtCore.QRect(460, 70, 161, 51))
                self.nome2.setStyleSheet("font: 16pt \"Largo\";  background-color:#FFBF46; color:#474044;     border-width: 2px;\n"
        "    border-radius: 10px;")
                self.nome2.setObjectName("nome2")
                
                self.nome3 = QtWidgets.QPushButton(self.centralwidget)
                self.nome3.setGeometry(QtCore.QRect(460, 130, 161, 51))
                self.nome3.setStyleSheet("font: 16pt \"Largo\";  background-color:#FFBF46; color:#474044;     border-width: 2px;\n"
        "    border-radius: 10px;")
                self.nome3.setObjectName("nome3")
                
                self.linkcontribute = QtWidgets.QPushButton(self.centralwidget)
                self.linkcontribute.setGeometry(QtCore.QRect(500, 200, 75, 51))
                self.linkcontribute.clicked.connect(link)
                self.linkcontribute.setStyleSheet("    border-width: 2px;\n"
        "    border-radius: 10px; background-color: #FFBF46")
                self.linkcontribute.setIcon(QIcon('github.png'))
                self.linkcontribute.setIconSize(QtCore.QSize(50, 50)) 
                self.linkcontribute.setObjectName("linkcontribute")
                
                Window.setCentralWidget(self.centralwidget)
                
                self.menubar = QtWidgets.QMenuBar(Window)
                self.menubar.setGeometry(QtCore.QRect(0, 0, 650, 21))
                self.menubar.setObjectName("menubar")
                
                Window.setMenuBar(self.menubar)
                
                self.statusbar = QtWidgets.QStatusBar(Window)
                self.statusbar.setObjectName("statusbar")
                
                Window.setStatusBar(self.statusbar)

                self.retranslateUi(Window)
                QtCore.QMetaObject.connectSlotsByName(Window)

        def retranslateUi(self, Window):
                _translate = QtCore.QCoreApplication.translate
                Window.setWindowTitle(_translate("Window", "Photoshop Text Editor"))
                self.sair.setText(_translate("Window", "SAIR"))
                self.salvarpng.setText(_translate("Window", "Salvar como .png"))
                self.inserir.setText(_translate("Window", "Inserir"))
                self.nome1.setText(_translate("Window", "NOME <"))
                self.nome2.setText(_translate("Window", "NOME >"))
                self.nome3.setText(_translate("Window", "NOME CIMA "))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Window = QtWidgets.QMainWindow()
    ui = Ui_Window()
    ui.setupUi(Window)
    Window.show()
    sys.exit(app.exec())
