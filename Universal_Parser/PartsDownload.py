from selenium.webdriver import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium import webdriver
import time
import pandas as pd
from tqdm import tqdm

#класс описывающий работу всех функций для выгрузки отдельных полей. Выгрузка происходит по идентификационным номерам.
class Parts:

    def __init__(self):
        pass

    #Тема вопроса
    def theme(self, driver):

        print('Выгрузка Тем началась:')

        docs_df = pd.read_excel("для_поиска_по_номеру.xlsx")
        themes = []
        doc_nums = []

        for doc_number in tqdm(docs_df["Номер обращения"]):
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                wait = WebDriverWait(driver, 1)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
                search_field = wait.until(EC.visibility_of_element_located((By.NAME, 'search')))
                search_field.send_keys(doc_number)
                press_search = driver.find_element(By.XPATH,
                                               '/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/button')
                press_search.click()
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                link_to_doc = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@name='s-doc__item']"))
                )
                link_to_doc.click()
                time.sleep(0.5)
            except:
                themes.append("Битый номер")
                doc_nums.append(doc_number)

            themes_counter = 0
            local_theme_list = []
            try:
                time.sleep(0.5)
                all_themes = driver.find_elements(By.XPATH, ".//*[@data-tour='72'][ @class='og-question-ref__value b3']/b")
                for item in all_themes:
                    themes_counter += 1
                    theme_parts = item.get_attribute("textContent")
                    theme_parts = theme_parts.replace('\n', ' ')
                    while "  " in theme_parts:
                        theme_parts = theme_parts.replace("  ", " ")
                    theme_parts = theme_parts.strip()
                    local_theme_list.append(theme_parts)
            except:
                local_theme_list.append("Тема не указана")

            themes.append(local_theme_list)
            doc_nums.append(doc_number)

        result_df = pd.DataFrame({'Номер обращения': doc_nums, 'Тема': themes})
        result_df.to_excel('Результаты/темы.xlsx')

        return result_df

    #Объект вопроса
    def query_obj(self, driver):

        print('Выгрузка Объектов вопросов началась:')

        docs_df = pd.read_excel("для_поиска_по_номеру.xlsx")
        objs = []
        doc_nums = []

        for doc_number in tqdm(docs_df["Номер обращения"]):
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                wait = WebDriverWait(driver, 1)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
                search_field = wait.until(EC.visibility_of_element_located((By.NAME, 'search')))
                search_field.send_keys(doc_number)
                press_search = driver.find_element(By.XPATH,
                                                        '/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/button')
                press_search.click()
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                link_to_doc = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@name='s-doc__item']"))
                )
                link_to_doc.click()
                time.sleep(0.5)
            except:
                objs.append("Битый номер")
                doc_nums.append(doc_number)

            local_org_list = []
            try:
                all_orgs = driver.find_elements(By.XPATH, ".//*[@class='block_b6_head'][@colspan=4]")

                for item in all_orgs:
                    organization_call = item.get_attribute("textContent")
                    organization_call = organization_call.replace('\n', ' ')
                    organization_call = organization_call.replace('Объект вопроса:', '')
                    while "  " in organization_call:
                        organization_call = organization_call.replace("  ", " ")
                    organization_call = organization_call.strip()
                    local_org_list.append(organization_call)
            except:
                local_org_list.append("МО не указана")

            objs.append(local_org_list)
            doc_nums.append(doc_number)

        result_df = pd.DataFrame({'Номер обращения': doc_nums, 'Объект': objs})
        result_df.to_excel('Результаты/объекты.xlsx')

        return result_df

    #Тип объекта вопроса
    def query_obj_type(self, driver):

        print('Выгрузка Типов объектов вопросов началась:')

        docs_df = pd.read_excel("для_поиска_по_номеру.xlsx")
        objs_type = []
        doc_nums = []

        for doc_number in tqdm(docs_df["Номер обращения"]):
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                wait = WebDriverWait(driver, 1)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
                search_field = wait.until(EC.visibility_of_element_located((By.NAME, 'search')))
                search_field.send_keys(doc_number)
                press_search = driver.find_element(By.XPATH,
                                                        '/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/button')
                press_search.click()
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                link_to_doc = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@name='s-doc__item']"))
                )
                link_to_doc.click()
                time.sleep(0.5)
            except:
                objs_type.append("Битый номер")
                doc_nums.append(doc_number)

            local_typ_list = []
            all_orgs = driver.find_elements(By.XPATH, ".//*[@class='block_b6_head'][@colspan=4]")
            all_typs = driver.find_elements(By.XPATH, ".//td[@class='b6']")
            for i in range(1, 10*len(all_orgs), 10):
                try:
                    if all_orgs[i%10-1].get_attribute("textContent") == 'без объекта вопроса':
                        local_typ_list.append('Тип не указан')
                    typ = all_typs[i].get_attribute("textContent")
                    typ = typ.replace('\n', ' ')
                    while "  " in typ:
                        typ = typ.replace("  ", " ")
                    typ = typ.strip()
                    local_typ_list.append(typ)
                except:
                    local_typ_list.append("Тип не указан")

                local_typ_list = local_typ_list[:len(all_orgs)]
                objs_type.append(local_typ_list)
                doc_nums.append(doc_number)

        result_df = pd.DataFrame({'Номер обращения': doc_nums, 'Тип объекта': objs_type})
        result_df.to_excel('Результаты/типы_объектов.xlsx')

        return result_df

    #Аннотация
    def query_annot(self, driver):

        print('Выгрузка Аннотаций началась:')

        docs_df = pd.read_excel("для_поиска_по_номеру.xlsx")
        annots = []
        doc_nums = []

        for doc_number in tqdm(docs_df["Номер обращения"]):
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                wait = WebDriverWait(driver, 1)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
                search_field = wait.until(EC.visibility_of_element_located((By.NAME, 'search')))
                search_field.send_keys(doc_number)
                press_search = driver.find_element(By.XPATH,
                                                        '/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/button')
                press_search.click()
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                link_to_doc = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@name='s-doc__item']"))
                )
                link_to_doc.click()
                time.sleep(0.5)
            except:
                annots.append("Битый номер")
                doc_nums.append(doc_number)

            local_annot_list = []
            try:
                all_annots = driver.find_elements(By.XPATH, ".//*[@class='og-question-ref__value b3'][@data-tour='73']")
                for item in all_annots:
                    annot = item.get_attribute("textContent")
                    annot = annot.replace('\n', ' ')
                    while "  " in annot:
                        annot = annot.replace("  ", " ")
                    annot = annot.strip()
                    local_annot_list.append(annot)
            except:
                local_annot_list.append("Аннотация не указана")

            annots.append(local_annot_list)
            doc_nums.append(doc_number)

            result_df = pd.DataFrame({'Номер обращения': doc_nums, 'Аннотация': annots})
            result_df.to_excel('Результаты/аннотации.xlsx')

        return result_df

    #Организация отправления
    def appeal_source(self, driver):

        print('Выгрузка Организации-отправителя началась:')

        docs_df = pd.read_excel("для_поиска_по_номеру.xlsx")
        sources = []
        doc_nums = []

        for doc_number in tqdm(list(docs_df["Номер обращения"])[1500:]):
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                wait = WebDriverWait(driver, 1)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
                search_field = wait.until(EC.visibility_of_element_located((By.NAME, 'search')))
                search_field.send_keys(doc_number)
                press_search = driver.find_element(By.XPATH,
                                                        '/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/button')
                press_search.click()
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                link_to_doc = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@name='s-doc__item']"))
                )
                link_to_doc.click()
                time.sleep(0.5)
            except:
                sources.append("Битый номер")
                doc_nums.append(doc_number)

            try:
                source = driver.find_element(By.XPATH, ".//*[@id='td_cl_sign']").get_attribute("textContent")
                sources.append(source)
                doc_nums.append(doc_number)
            except:
                sources.append("Не найдено")
                doc_nums.append(doc_number)

            if len(sources)%500==0:
                part_res = pd.DataFrame({'Номер обращения': doc_nums, 'Отправитель': sources})
                part_res.to_excel('Текущий_прогресс_2.xlsx')

        result_df = pd.DataFrame({'Номер обращения': doc_nums, 'Отправитель': sources})
        result_df.to_excel('Результаты/отправители.xlsx')

        return result_df

    #Кол-во подписей
    def sign_counts(self, driver):

        print('Выгрузка Кол-ва подписей обращений началась:')

        docs_df = pd.read_excel("для_поиска_по_номеру.xlsx")
        signs = []
        doc_nums = []

        for doc_number in tqdm(list(docs_df["Номер обращения"])):
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                wait = WebDriverWait(driver, 1)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
                search_field = wait.until(EC.visibility_of_element_located((By.NAME, 'search')))
                search_field.send_keys(doc_number)
                press_search = driver.find_element(By.XPATH,
                                                        '/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/button')
                press_search.click()
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                link_to_doc = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@name='s-doc__item']"))
                )
                link_to_doc.click()
                time.sleep(0.5)
            except:
                signs.append("Битый номер")
                doc_nums.append(doc_number)

            try:
                sign = driver.find_element(By.XPATH, "//*[text()='Кол-во подписей:']/following-sibling::td").get_attribute("textContent").strip()
                signs.append(sign)
                doc_nums.append(doc_number)
            except:
                signs.append("Не найдено")
                doc_nums.append(doc_number)

        result_df = pd.DataFrame({'Номер обращения': doc_nums, 'Кол-во подписей': signs})
        result_df.to_excel('Результаты/подписи.xlsx')

        return result_df

    #Дата поступления обращения
    def income_date(self, driver):

        print('Выгрузка Даты поступления обращения началась:')

        docs_df = pd.read_excel("для_поиска_по_номеру.xlsx")
        dates = []
        doc_nums = []

        for doc_number in tqdm(list(docs_df["Номер обращения"])):
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                wait = WebDriverWait(driver, 1)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
                search_field = wait.until(EC.visibility_of_element_located((By.NAME, 'search')))
                search_field.send_keys(doc_number)
                press_search = driver.find_element(By.XPATH,
                                                        '/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/button')
                press_search.click()
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                link_to_doc = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@name='s-doc__item']"))
                )
                link_to_doc.click()
                time.sleep(0.5)
            except:
                dates.append("Битый номер")
                doc_nums.append(doc_number)

            try:
                date = driver.find_element(By.XPATH, "//*[@data-tour='48'][@class='b3']").get_attribute("textContent").strip()
                dates.append(date)
                doc_nums.append(doc_number)
            except:
                dates.append("Не найдено")
                doc_nums.append(doc_number)

        result_df = pd.DataFrame({'Номер обращения': doc_nums, 'Дата поступления': dates})
        result_df.to_excel('Результаты/даты_поступления.xlsx')

        return result_df

    #Дата регистрации обращения
    def registartion_date(self, driver):

        print('Выгрузка Даты регистрации обращения началась:')

        docs_df = pd.read_excel("для_поиска_по_номеру.xlsx")
        dates = []
        doc_nums = []

        for doc_number in tqdm(list(docs_df["Номер обращения"])):
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                wait = WebDriverWait(driver, 1)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
                search_field = wait.until(EC.visibility_of_element_located((By.NAME, 'search')))
                search_field.send_keys(doc_number)
                press_search = driver.find_element(By.XPATH,
                                                        '/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/button')
                press_search.click()
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                link_to_doc = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@name='s-doc__item']"))
                )
                link_to_doc.click()
                time.sleep(0.5)
            except:
                dates.append("Битый номер")
                doc_nums.append(doc_number)

            try:
                date = driver.find_element(By.XPATH, "//*[@class='b fixedWidth']/span").get_attribute("textContent").strip()
                dates.append(date)
                doc_nums.append(doc_number)
            except:
                dates.append("Не найдено")
                doc_nums.append(doc_number)

        result_df = pd.DataFrame({'Номер обращения': doc_nums, 'Дата регистрации': dates})
        result_df.to_excel('Результаты/даты_регистрации.xlsx')

        return result_df

    #ФИО автора обращения
    def author_fio(self, driver):

        print('Выгрузка ФИО авторов началась:')

        docs_df = pd.read_excel("для_поиска_по_номеру.xlsx")
        fios = []
        doc_nums = []

        for doc_number in tqdm(list(docs_df["Номер обращения"])):
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                wait = WebDriverWait(driver, 1)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
                search_field = wait.until(EC.visibility_of_element_located((By.NAME, 'search')))
                search_field.send_keys(doc_number)
                press_search = driver.find_element(By.XPATH,
                                                        '/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/button')
                press_search.click()
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                link_to_doc = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@name='s-doc__item']"))
                )
                link_to_doc.click()
                time.sleep(0.5)
            except:
                fios.append("Битый номер")
                doc_nums.append(doc_number)

            try:
                fio = driver.find_element(By.XPATH, "//*[text()='Автор обращения:']/following-sibling::td").get_attribute("textContent").strip()
                fios.append(fio)
                doc_nums.append(doc_number)
            except:
                fios.append("Не найдено")
                doc_nums.append(doc_number)

        result_df = pd.DataFrame({'Номер обращения': doc_nums, 'ФИО автора': fios})
        result_df.to_excel('Результаты/фио_автора.xlsx')

        return result_df

    #Почта автора
    def author_mail(self, driver):

        print('Выгрузка Почта авторов началась:')

        docs_df = pd.read_excel("для_поиска_по_номеру.xlsx")
        mails = []
        doc_nums = []

        for doc_number in tqdm(list(docs_df["Номер обращения"])):
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                wait = WebDriverWait(driver, 1)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
                search_field = wait.until(EC.visibility_of_element_located((By.NAME, 'search')))
                search_field.send_keys(doc_number)
                press_search = driver.find_element(By.XPATH,
                                                        '/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/button')
                press_search.click()
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                link_to_doc = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@name='s-doc__item']"))
                )
                link_to_doc.click()
                time.sleep(1)
            except:
                mails.append("Битый номер")
                doc_nums.append(doc_number)

            try:
                mail = driver.find_element(By.XPATH, "//*[text()='E-mail:']/following-sibling::td").get_attribute("textContent").strip()
                mails.append(mail)
                doc_nums.append(doc_number)
            except:
                mails.append("Не найдено")
                doc_nums.append(doc_number)

        result_df = pd.DataFrame({'Номер обращения': doc_nums, 'Почта': mails})
        result_df.to_excel('Результаты/почта_автора.xlsx')

        return result_df


    #Текст обращения
    def txt(self, driver):

        print('Выгрузка Текстов обращений началась:')

        docs_df = pd.read_excel("для_поиска_по_номеру.xlsx")
        txts = []
        doc_nums = []

        for doc_number in tqdm(list(docs_df["Номер обращения"])):
            try:
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)
                wait = WebDriverWait(driver, 1)
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.F2)
                search_field = wait.until(EC.visibility_of_element_located((By.NAME, 'search')))
                search_field.send_keys(doc_number)
                press_search = driver.find_element(By.XPATH,
                                                        '/html/body/div[5]/div[2]/div/div/div[2]/div/div[2]/div[1]/button')
                press_search.click()
                driver.find_element(By.TAG_NAME, 'body').send_keys(Keys.ESCAPE)

                link_to_doc = wait.until(
                    EC.visibility_of_element_located(
                        (By.XPATH, "//*[@name='s-doc__item']"))
                )
                link_to_doc.click()
                time.sleep(1)

                txt_doc_type = wait.until(
                    EC.visibility_of_element_located((By.XPATH, ".//*[@class='s-viewer__mode-button']"))
                )
                txt_doc_type.click()
            except:
                txts.append("Битый номер")
                doc_nums.append(doc_number)

            txt_seq = ""
            try:
                more_pages = wait.until(
                    EC.visibility_of_element_located((By.XPATH, ".//*[@class='show_all_pages'][1]"))
                )
                more_pages.click()

                time.sleep(2)
                pages = driver.find_elements(By.XPATH, ".//*[@class='page-layout__body highlightable']")
                #print(len(pages))

                count = 0
                for page in pages:
                    if count < 5:
                        txt_seq = txt_seq + '//' + page.text
                        count += 1
                    else:
                        break

            except:
                txt_seq = driver.find_element(By.XPATH, f".//*[@class='page-layout__body highlightable']").text

            txts.append(txt_seq)
            doc_nums.append(doc_number)

        result_df = pd.DataFrame({'Номер обращения': doc_nums, 'Текст': txts})
        result_df.to_excel('Результаты/почта_автора.xlsx')

        return result_df