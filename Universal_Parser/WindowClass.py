import logging

from tqdm import tqdm
import sys
import pandas as pd
pd.options.mode.chained_assignment = None
from selenium.webdriver import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import date
import time
import os
import shutil
from PyQt5.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDateTimeEdit,
    QDial,
    QDoubleSpinBox,
    QFontComboBox,
    QLabel,
    QLCDNumber,
    QLineEdit,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSlider,
    QSpinBox,
    QTimeEdit,
    QVBoxLayout,
    QHBoxLayout,
    QWidget,
    QGridLayout,
    QFrame,
)
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager


#класс, инициализирующий пользовательское окно с чекбоксами для выбора необходимых полей для выгрузки.
class MainChoiceWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        #описание внешнего вида и обработки действий пользователя (изменений состояний чекбоксов, нажатие на кнопки)
        self.centralwidget = QWidget()
        self.centralwidget.setObjectName("centralwidget")
        self.centralwidget.setStyleSheet("background-color: rgb(0, 0, 0);")
        self.setCentralWidget(self.centralwidget)

        self.checkBoxes = {
            'Тема вопроса': '', 'Объект вопроса': '', 'Тип объекта вопроса': '', 'Аннотация': '', 'Организация отправления': '', 'Кол-во подписей': '',
            'ФИО автора обращения': '', 'Почта автора': '', 'Дата поступления обращения': '', 'Дата регистрации обращения': '', 'Текст обращения': '',
        }

        column_cnt = 2
        row = 0
        layout = QGridLayout(self.centralwidget)
        self.check_boxes = []

        for step, name in enumerate(self.checkBoxes):
            cb = QCheckBox(name, objectName=name.lower())
            cb.setStyleSheet("background-color: rgb(0, 0, 0);""color: rgb(0, 150, 150);")
            row = step // column_cnt
            col = step % column_cnt
            layout.addWidget(cb, row, col)
            self.check_boxes.append(cb)

        self.download_button = QPushButton('Выгрузить')
        self.download_button.setStyleSheet("background-color: rgb(100, 100, 100);""color: rgb(150, 150, 0);")
        self.exit_button = QPushButton('Выход')
        self.exit_button.setStyleSheet("background-color: rgb(100, 100, 100);""color: rgb(150, 150, 0);")

        self.frame = QFrame()
        layoutV = QVBoxLayout(self.frame)
        layoutV.setContentsMargins(0, 0, 0, 0)
        layoutV.addWidget(self.download_button)
        layoutV.addWidget(self.exit_button)

        layout.addWidget(self.frame, row+1, 0, 1, column_cnt)

        self.download_button.clicked.connect(self.plan_maker)
        self.exit_button.clicked.connect(self.exit_click)

    #функция, вызываемая нажатием кнопки, и собирающая состояния чекбоксов. Формируется список состояний, который используется в классе DownLoadClass
    def plan_maker(self):
        self.items_to_download = []
        for box in self.check_boxes:
            if box.isChecked():
                item = box.text()
                self.items_to_download.append(item)

        print("Скачать: ", self.items_to_download)

    def exit_click(self):
        QApplication.quit()



if __name__ == '__main__':
    app = QApplication(sys.argv)

    w = MainChoiceWindow()
    w.show()

    sys.exit(app.exec_())