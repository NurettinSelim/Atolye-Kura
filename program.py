import sys

from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QMovie
from PyQt5.QtWidgets import QFileDialog

import draw_students
from random import shuffle


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi("kendin_yap_arayuz.ui", self)
        self.gif = QMovie('gif.gif')
        self.upload_button = self.findChild(QtWidgets.QPushButton, 'uploadButton')
        self.upload_button.clicked.connect(self.upload_list)
        self.upload_button = self.findChild(QtWidgets.QPushButton, 'drawButton')
        self.upload_button.clicked.connect(self.draw_process)

        self.gif_label = self.findChild(QtWidgets.QLabel, 'gifLabel')
        self.text_label = self.findChild(QtWidgets.QLabel, 'textLabel')

        self.list_widget = self.findChild(QtWidgets.QListWidget, "listWidget")

        # Sınıf İsimleri
        self.text_dict = dict()
        self.text_sinif3 = self.findChild(QtWidgets.QLabel, 'textSinif3')
        self.text_dict["3"] = self.text_sinif3
        self.text_sinif4 = self.findChild(QtWidgets.QLabel, 'textSinif4')
        self.text_dict["4"] = self.text_sinif4
        self.text_sinif5 = self.findChild(QtWidgets.QLabel, 'textSinif5')
        self.text_dict["5"] = self.text_sinif5
        self.text_sinif6 = self.findChild(QtWidgets.QLabel, 'textSinif6')
        self.text_dict["6"] = self.text_sinif6
        self.text_sinif7 = self.findChild(QtWidgets.QLabel, 'textSinif7')
        self.text_dict["7"] = self.text_sinif7

        # Sınıf Listeleri
        self.liste_dict = dict()
        self.liste3 = self.findChild(QtWidgets.QListWidget, 'liste3')
        self.liste_dict["3"] = self.liste3
        self.liste4 = self.findChild(QtWidgets.QListWidget, 'liste4')
        self.liste_dict["4"] = self.liste4
        self.liste5 = self.findChild(QtWidgets.QListWidget, 'liste5')
        self.liste_dict["5"] = self.liste5
        self.liste6 = self.findChild(QtWidgets.QListWidget, 'liste6')
        self.liste_dict["6"] = self.liste6
        self.liste7 = self.findChild(QtWidgets.QListWidget, 'liste7')
        self.liste_dict["7"] = self.liste7

        for i in self.text_dict.values():
            i.setVisible(False)
        for i in self.liste_dict.values():
            i.setVisible(False)

        self.filename = str()
        self.counter = 3
        self.grade_list = dict()
        self.show()

    def upload_list(self):
        self.filename, _ = QFileDialog.getOpenFileName(filter="Excell Dosyası (*.xlsx)")
        if self.filename:
            self.text_label.setText("Dosya yüklemesi tamamlandı.")

    def draw_process(self):
        self.gif_label.setMovie(self.gif)
        self.gif.start()
        QTimer.singleShot(2000, self.draw_to_list)

    def draw_to_list(self):
        self.gif_label.clear()
        students_list = draw_students.draw_students_from_excel(self.filename)
        self.grade_list = draw_students.create_grade_list(students_list)
        draw_students.save_to_excel(self.grade_list)
        self.create_list_view()

    def create_list_view(self):
        if self.counter <= 7:
            self.text_dict[str(self.counter)].setVisible(True)
            self.liste_dict[str(self.counter)].setVisible(True)
            shuffle(self.grade_list[str(self.counter)])
            for student in self.grade_list[str(self.counter)]:
                self.liste_dict[str(self.counter)].addItem(student.to_text())
            self.counter += 1
            QTimer.singleShot(4000, self.create_list_view)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = Ui()
    app.exec_()
