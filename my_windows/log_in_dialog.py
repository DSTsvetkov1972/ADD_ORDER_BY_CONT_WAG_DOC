# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'log_in_dialog.ui'
##
## Created by: Qt User Interface Compiler version 6.6.3
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform, Qt)
from PySide6.QtWidgets import (QAbstractButton, QApplication, QDialog, QDialogButtonBox,
    QFormLayout, QLineEdit, QTextEdit,QSizePolicy, QWidget)


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        
        Dialog.setFixedSize(500, 300)
        Dialog.setWindowTitle("Введите параметры подключения")
        Dialog.setWindowModality(Qt.WindowModality.ApplicationModal)       

        self.buttonBox = QDialogButtonBox(Dialog)
        self.buttonBox.setObjectName(u"buttonBox")
        self.buttonBox.setGeometry(QRect(30, 240, 441, 32))
        self.buttonBox.setOrientation(Qt.Horizontal)
        #self.buttonBox.setStandardButtons(QDialogButtonBox.Cancel|QDialogButtonBox.Ok)
        self.buttonBox.setStandardButtons(QDialogButtonBox.Ok)

        self.formLayoutWidget = QWidget(Dialog)
        self.formLayoutWidget.setGeometry(QRect(10, 10, 481, 161))
        self.formLayout = QFormLayout(self.formLayoutWidget)
        self.formLayout.setContentsMargins(0, 0, 0, 0)

        self.lineEditHostLabel = QLineEdit(self.formLayoutWidget)
        self.lineEditHostLabel.setEnabled(False)
        self.lineEditHostLabel.setStyleSheet('border: none; background-color: rgba(0, 0, 0, 0.0); font-weight: 700')
        self.lineEditHostLabel.setText("HOST")
        self.lineEditHostLabel.setAlignment(Qt.AlignRight)
        self.formLayout.setWidget(0, QFormLayout.LabelRole, self.lineEditHostLabel)

        self.lineEditHostField = QLineEdit(self.formLayoutWidget)
        self.formLayout.setWidget(0, QFormLayout.FieldRole, self.lineEditHostField)

        self.lineEditPortLabel = QLineEdit(self.formLayoutWidget)
        self.lineEditPortLabel.setEnabled(False)
        self.lineEditPortLabel.setStyleSheet('border: none; background-color: rgba(0, 0, 0, 0.0); font-weight: 700')
        self.lineEditPortLabel.setText("PORT")
        self.lineEditPortLabel.setAlignment(Qt.AlignRight)        
        self.formLayout.setWidget(1, QFormLayout.LabelRole, self.lineEditPortLabel)

        self.lineEditPortField = QLineEdit(self.formLayoutWidget)
        self.formLayout.setWidget(1, QFormLayout.FieldRole, self.lineEditPortField)

        self.lineEditDBNameLabel = QLineEdit(self.formLayoutWidget)
        self.lineEditDBNameLabel.setEnabled(False)
        self.lineEditDBNameLabel.setStyleSheet('border: none; background-color: rgba(0, 0, 0, 0.0); font-weight: 700')
        self.lineEditDBNameLabel.setText("DBNAME")
        self.lineEditDBNameLabel.setAlignment(Qt.AlignRight)
        self.formLayout.setWidget(2, QFormLayout.LabelRole, self.lineEditDBNameLabel)

        self.lineEditDBNameField = QLineEdit(self.formLayoutWidget)
        self.formLayout.setWidget(2, QFormLayout.FieldRole, self.lineEditDBNameField)

        self.lineEditUserLabel = QLineEdit(self.formLayoutWidget)
        self.lineEditUserLabel.setEnabled(False)
        self.lineEditUserLabel.setStyleSheet('border: none; background-color: rgba(0, 0, 0, 0.0); font-weight: 700')
        self.lineEditUserLabel.setText("USER")
        self.lineEditUserLabel.setAlignment(Qt.AlignRight)        
        self.formLayout.setWidget(3, QFormLayout.LabelRole, self.lineEditUserLabel)

        self.lineEditUserField = QLineEdit(self.formLayoutWidget)
        self.formLayout.setWidget(3, QFormLayout.FieldRole, self.lineEditUserField)


        self.lineEditPasswordLabel = QLineEdit(self.formLayoutWidget)
        self.lineEditPasswordLabel.setEnabled(False)
        self.lineEditPasswordLabel.setStyleSheet('border: none; background-color: rgba(0, 0, 0, 0.0); font-weight: 700')
        self.lineEditPasswordLabel.setText("PASSWORD")
        self.lineEditPasswordLabel.setAlignment(Qt.AlignRight)        
        self.formLayout.setWidget(4, QFormLayout.LabelRole, self.lineEditPasswordLabel)

        self.lineEditPasswordField = QLineEdit(self.formLayoutWidget)
        self.lineEditPasswordField.setEchoMode(QLineEdit.Password)
        self.formLayout.setWidget(4, QFormLayout.FieldRole, self.lineEditPasswordField)

        
        #self.retranslateUi(Dialog)
        #self
        #self.buttonBox.rejected.connect(Dialog.reject)

        #QMetaObject.connectSlotsByName(Dialog)

    #def accept (self):
    #    print("dffffff")
        

    # setupUi

    #def retranslateUi(self, Dialog):
    #    Dialog.setWindowTitle(QCoreApplication.translate("Dialog", u"Dialog", None))

        #self.lineEdit_2.setText("")
        #self.lineEdit_4.setText(QCoreApplication.translate("Dialog", u"PORT", None))
        #self.lineEdit_3.setText(QCoreApplication.translate("Dialog", u"DBNAME", None))
        #self.lineEdit_5.setText(QCoreApplication.translate("Dialog", u"USER", None))
        #self.lineEdit_7.setText(QCoreApplication.translate("Dialog", u"PASSWORD", None))
    # retranslateUi


################################################################################ 

class LogInDialog(QWidget):
    def __init__ (self, parent=None):
        QWidget.__init__(self, parent)
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)

        self.ui.buttonBox.accepted.connect(self.accept)
        self.ui.lineEditHostField.textChanged.connect(lambda: print('textChanged'))    
          
    def accept(self):

        print(self.ui.lineEditHostField.text())
        #self.close()

  


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = LogInDialog()
    window.show()
    #w = LogInDialog()
    #w.show()
    sys.exit(app.exec())    