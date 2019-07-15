from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import *
import unittest
import time
import datetime
import logging
from dominios import get_domains_list
from dominios import path_screenshots

logging.basicConfig(filename='registro.log', level=logging.INFO,
    format='%(asctime)s: %(levelname)s: %(message)s',
    datefmt="%d/%m/%Y %I:%M:%S")

user_email = "tu_email@mail.com"
user_pass = "tu_password"
url_domain = "https://www.mlcatalog.com/catalog/apd-frontend/domains/"


class Planillas(unittest.TestCase):
    def setUp(self):
        self.inicio = time.time()
        self.driver = webdriver.Chrome()
        # self.driver = webdriver.Remote(command_executor="http://34.206.108.180:4444/wd/hub",
        # desired_capabilities={"browserName": "chrome", "javascriptEnabled": True})
        self.driver.implicitly_wait(1000)
        self.wait = WebDriverWait(self.driver, 1000)
        self.driver.get("https://accounts.google.com/signin/v2/identifier?continue=https%3A%2F%2Fmail.google.com"
                        "%2Fmail%2F&service=mail&sacu=1&rip=1&flowName=GlifWebSignIn&flowEntry=ServiceLogin")
        self.driver.maximize_window()
        time.sleep(2)

    def test_login_user(self):
        comienzo_test = str(datetime.datetime.now().time())
        try:
            self.user_login()
            lista_dominios = get_domains_list()
            for dominio in lista_dominios:
                self.open_catalog_domain(dominio)
                self.generate_sheet(dominio)
                self.click_product(dominio)
                self.save_file(dominio)
                self.driver.find_element_by_tag_name('body').send_keys(Keys.CONTROL + 'w')
        except Exception as e:
            logging.error(e, '¡ERROR! - Hubo un error en la ejecucion')
        logging.info("Tiempo de inicio --> " + comienzo_test)
        logging.info("Tiempo de fin --> " + str(datetime.datetime.now().time()))
        final = time.time()
        logging.info("TIEMPO DE EJECUCION --> " + str((final - self.inicio) / 60))
        logging.info("============================================================")

    def user_login(self):
        """ Loguea al email de nahualit"""
        try:
            input_email = self.driver.find_element_by_xpath("//input[@type='email']")
            input_email.send_keys(user_email)
            logging.info("Login: Email")
            self.driver.find_element_by_xpath("//span[.='Siguiente' or .='Next']").click()
            input_password = self.driver.find_element_by_name("password")
            input_password.send_keys(user_pass)
            logging.info("login: password")
            self.wait.until(ec.element_to_be_clickable((By.XPATH, "//span[.='Siguiente' or .='Next']"))).click()
            logging.info("Login: Siguiente")
            time.sleep(3)
        except Exception as e:
            self.take_screenshot("login")
            logging.error(e, "----- ¡ERROR! - No se pudo iniciar sesion -----")

    def open_catalog_domain(self, domain):
        """ Abre el catalogo del dominio enviado por parametro"""
        try:
            self.driver.get(url_domain + domain + "/products")
            self.wait.until(ec.element_to_be_clickable((By.XPATH, '//button[.="Generar planilla View Only"]')))
            dominio = domain.lower()
            self.wait.until(ec.text_to_be_present_in_element((By.TAG_NAME, 'H1'), 'Products ' + dominio))
            logging.info("Se ingreso al catalogo del dominio --> " + domain)
        except Exception as e:
            logging.error(e, "¡ERROR! Dominio no encontrado")
            self.take_screenshot(domain)

    def generate_sheet(self, dominio):
        """ Genera la planilla view only"""
        self.driver.find_element_by_xpath('//button[.="Generar planilla View Only"]').click()  # click en el boton
        logging.info("Click en boton 'Generar planilla view only' ")
        try:
            cargando = self.driver.find_element_by_class_name("loading")
            if cargando:
                logging.info("Generando planilla view only...")
            link = self.driver.find_element_by_xpath('//p[@id="modal-text"]/a').get_attribute("href")
            self.driver.get(link)
            logging.info("PLANILLA GENERADA")
        except Exception as e:
            self.take_screenshot(dominio)
            logging.error(e, "----- ERROR AL GENERAR LA PLANILLA -----")
            logging.error(self.driver.current_url)

    def click_product(self, dominio):
        """ Selecciona en la tabla el dominio enviado por parametro """
        try:
            self.wait.until(ec.presence_of_element_located((By.XPATH, '//td/a')))
            sheet = self.driver.find_element_by_xpath('//td/a').get_attribute('href')
            self.driver.get(sheet)
            self.driver.switch_to.window(self.driver.window_handles[1])
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[0])
        except Exception as e:
            self.take_screenshot(dominio)
            logging.info(e, "¡ERROR! No se pudo hacer click en el producto")

    def save_file(self, dominio):
        """ Descarga la planilla Spreedsheet en formato xlsx"""
        try:
            boton_compartir = (By.XPATH, "//div[@role='button' and text()='Compartir' or text()='Share']")
            self.wait.until(ec.presence_of_element_located(boton_compartir))
            button_file = self.driver.find_element_by_id("docs-file-menu")
            time.sleep(3)
            button_file.click()
            logging.info("Click en 'Archivo' ")
            save_as = self.driver.find_element_by_xpath(
                "//span[@aria-label='Descargar como d' or @aria-label='Download as d']")
            save_as.click()
            logging.info("Click en 'Guardar como' ")
            file_excel = self.driver.find_element_by_xpath("//span[@aria-label='Microsoft Excel (.xlsx) x']")
            file_excel.click()
            logging.info("Guardando como Excel: " + dominio)
            time.sleep(5)
        except Exception as e:
            logging.error(e, "¡ERROR! No se pudo guardar el archivo")

    def take_screenshot(self, dominio):
        """ Genera un captura de pantalla """
        filename = "/screenshot_" + dominio + str(round(time.time() * 1000)) + ".png"
        directory = path_screenshots()
        screenshot_directory = str(directory) + filename
        try:
            self.driver.save_screenshot(screenshot_directory)
            logging.info("Screeshot guardado en --> " + screenshot_directory)
        except FileNotFoundError:
            logging.error("No es un directorio valido :( " + str(directory))

    def tearDown(self):
        self.driver.quit()


if __name__ == '__main__':
    unittest.main()
