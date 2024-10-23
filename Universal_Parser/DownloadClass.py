import sys
import pandas as pd
pd.options.mode.chained_assignment = None
from selenium.webdriver import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import shutil
from PyQt5.QtWidgets import (
    QApplication,
)
from selenium import webdriver
from concurrent.futures import ThreadPoolExecutor
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from WindowClass import MainChoiceWindow
from PartsDownload import Parts


#данный класс служит для инициализации окон браузера для дальнейшего взаимодействия с элементами страниц  с помощью
#Selenium Webdriver. Запуск работы драйвера и обработка полученных сигналов из окна пользователя.
class DownLoadClass(MainChoiceWindow): #наследование от класса основного окна пользователя, с которого приходят сигналы

    def __init__(self):
        super().__init__()

        self.driver_1 = webdriver.Chrome()
        self.driver_2 = webdriver.Chrome()
        self.driver_list = [self.driver_1, self.driver_2]
        self.download_button.clicked.connect(self.driver_work)
        self.parts = Parts()

    def DownLoadInit(self):

        #перечень ключей, по которым происходит обращение и запуск функций, описанных в классе Parts, для выгрузки определенных пользователем, полей страниц
        dict_func_to_push = {
            'Тема вопроса': self.parts.theme,
            'Организация отправления': self.parts.appeal_source,
            'Аннотация': self.parts.query_annot,
            'Объект вопроса': self.parts.query_obj,
            'Тип объекта вопроса': self.parts.query_obj_type,
            'Кол-во подписей': self.parts.sign_counts,
            'ФИО автора обращения': self.parts.author_fio,
            'Почта автора': self.parts.author_mail,
            'Дата поступления обращения': self.parts.income_date,
            'Дата регистрации обращения': self.parts.registartion_date,
            'Текст обращения': self.parts.txt,
        }

        self.result_folder_refresh()

        items_for_window_one, items_for_window_two = self.items_to_download[:int((len(self.items_to_download)/2))], self.items_to_download[int((len(self.items_to_download)/2)):]

        #инициализация работ необходимых функций
        def downloading(lst, driver):
            for i in range(len(lst)):
                dict_func_to_push.get(lst[i])(driver)
                print(f'Выгружено: {lst[i]}')

        #добавление многопоточности для более быстрой выгрузки в случае множественного выбора полей для выгрузки
        with ThreadPoolExecutor() as executor:
            executor.submit(downloading, items_for_window_one, self.driver_1)
            executor.submit(downloading, items_for_window_two, self.driver_2)

        self.results_union()

    #запуск работы драйвера и прохождения всех этапов аутентификации (**** - скрытая ифнормация)
    def driver_work(self):
        while True:
            try:
                for driver in self.driver_list:
                    driver.get(
                        '****')

                    # Нахождение выпадающего списка организаций
                    organization_dropdown = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, 'organizations'))
                    )
                    organization_dropdown.click()
                    organization_dropdown.send_keys("****")
                    # Выбор определенной организации из выпадающего списка
                    desired_organization = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.LINK_TEXT, '****'))
                    )
                    desired_organization.click()
                    # Нахождение выпадающего списка фамилий
                    surname_dropdown = WebDriverWait(driver, 20).until(
                        EC.element_to_be_clickable((By.ID, 'logins'))
                    )
                    surname_dropdown.click()
                    surname_dropdown.send_keys("****")
                    # Выбор определенной фамилии из выпадающего списка
                    select_surname = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.LINK_TEXT, '****'))
                    )
                    select_surname.click()
                    # Ввод пароля
                    password_input = driver.find_element(By.ID, 'password_input')
                    password_input.send_keys('****')
                    # Нажатие кнопки отправки формы
                    submit_button = driver.find_element(By.XPATH,
                                                             '//*[@id="login_form"]/table/tbody/tr[11]/td[4]/input[1]')
                    submit_button.click()
                    driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                self.DownLoadInit()

                for driver in self.driver_list:
                    driver.quit()

                break

            except Exception as e:
                print(f"Произошла ошибка: {str(e)}")
                print("Повторное выполнение кода...")
                break

    #функция обработки результатов выгрузки
    def results_union(self):

        path = 'Результаты/'
        files = os.listdir(path)
        res_df = pd.read_excel(path + files[0], index_col=0)

        for file in files[1:]:
            buf_df = pd.read_excel(path + file, index_col=0)
            res_df = pd.concat([res_df, buf_df[buf_df.columns[1:]]], axis=1, ignore_index=False)

        res_df.to_excel(path+'Результаты_выгрузки.xlsx')
        print("Выгрузка готова!")

    def result_folder_refresh(self, folder_path='Результаты/'):
        shutil.rmtree(folder_path)
        os.mkdir(folder_path)




if __name__ == '__main__':
    app = QApplication(sys.argv)

    w = DownLoadClass()
    w.show()

    sys.exit(app.exec_())
