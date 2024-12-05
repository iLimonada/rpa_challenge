from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from services.excel_service import ExcelService
from config.settings import settings
import time

class RPAController:
    def __init__(self):
        self.excel_service = ExcelService(settings.EXCEL_FILE)

    def iniciar_planilha(self):
        try:
            dados = self.excel_service.extrair_dados()
            if not dados:
                print("Nada encontrado")
                return 0

            driver = webdriver.Chrome()
            driver.maximize_window()
            driver.get(settings.URL_RPA_CHALLENGE)

            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//a[@href='./assets/downloadFiles/challenge.xlsx']"))
                )
                print("Pagina carregada com sucesso")
            except Exception as ex:
                print(f"Erro ao carregar a página: {ex}")
                driver.quit()
                return 0


            botao_download_excel = driver.find_element(By.XPATH, "//a[@href='./assets/downloadFiles/challenge.xlsx']")
            botao_download_excel.click()

            botao_comecar = driver.find_element(By.XPATH, "//button[contains(text(), 'Start')]")
            count = 0
            while count < 1:
                botao_comecar.click()
                count += 1

            time.sleep(3)

            for lista_de_dados in dados:
                self.preencher_formulario(driver, lista_de_dados)

            time.sleep(5)
            driver.quit()

        except Exception as ex:
            print(f"Erro no programa {ex}")

    def preencher_formulario(self, driver, lista_de_dados):
        try:
            campo_primeiro_nome = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelFirstName"]')
            campo_primeiro_nome.send_keys(lista_de_dados[0])

            campo_ultimo_nome = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelLastName"]')
            campo_ultimo_nome.send_keys(lista_de_dados[1])

            campo_nome_empresa = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelCompanyName"]')
            campo_nome_empresa.send_keys(lista_de_dados[2])

            campo_cargo_empresa = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelRole"]')
            campo_cargo_empresa.send_keys(lista_de_dados[3])

            campo_endereco = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelAddress"]')
            campo_endereco.send_keys(lista_de_dados[4])

            campo_email = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelEmail"]')
            campo_email.send_keys(lista_de_dados[5])

            campo_telefone = driver.find_element(By.XPATH, '//input[@ng-reflect-name="labelPhone"]')
            campo_telefone.send_keys(lista_de_dados[6])

            submit_button = driver.find_element(By.CLASS_NAME, 'btn.uiColorButton')
            submit_button.click()

            time.sleep(1)
        except Exception as e:
            print(f"Erro ao preencher o formulário: {e}")
