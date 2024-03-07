from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
from time import sleep
import os
import threading
import shutil

options = Options()
options.add_experimental_option("detach", True)


download_directory = "C:\\Users\\redst\\Desktop\\DSL_PDF_EXPORT"


url = 'https://tools.oregonexplorer.info/OE_HtmlViewer/Index.html?viewer=renewable'

directory = "C:\\Users\\redst\\Downloads\\Zipped"
zip_list = []

for root, directories, files, in os.walk(directory):
        for file in files:
            zip_list.append(file)

class FileProcessingThread(threading.Thread):
    def __init__(self, currentIndex: int, jump: int):
        threading.Thread.__init__(self)
        self.currentIndex = currentIndex
        self.jump = jump
    
    def run(self):
        
        # while(self.currentIndex < zip_list_len):
        for i in range(1):
            workingName = str(zip_list[self.currentIndex]).removesuffix('.zip')
            chromeDriverManager = ChromeDriverManager()
            actual_download_path = os.path.join(download_directory,workingName)
            print(f"Download Path index: {self.currentIndex}, Path: {actual_download_path}")
            options.add_experimental_option('prefs', {
            "download.default_directory": actual_download_path, #Change default directory for downloads
            "download.prompt_for_download": False, #To auto download the file
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
            })
            os.mkdir(actual_download_path)
            driver = webdriver.Chrome(service=Service(chromeDriverManager.install()), 
                                options=options)
            
            driver.get(url)
            #driver.maximize_window()
            sleep(15)
            # close_splash = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, "panel-header-button right close-16 bound-visible")))
            """close_splash = driver.find_element(By.CLASS_NAME, "panel-header-button right close-16 bound-visible")
            close_splash.click()"""

            # clicking splash screen button
            #WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CLASS_NAME, "oeFormActionsRow")))
            driver.find_elements(By.CLASS_NAME, "oeFormActionsRow")[1].find_elements(By.TAG_NAME, "button")[1].click()
            sleep(1)

            # clicking disclaimer next button
            driver.find_element(By.CLASS_NAME, "ButtonBar.gcx-forms-footer").find_elements(By.TAG_NAME, "Button")[1].click()
            sleep(0.1)

            # filling out the site name & other info
            # site name
            driver.find_element(By.CLASS_NAME, "Body").find_element(By.TAG_NAME, "input").send_keys(workingName)
            sleep(0.1)

            # doing development checkboxes
            development_elements = driver.find_element(By.ID, "gcx-forms-checkGroupDevTypesOnShore-18").find_element(By.CLASS_NAME, "items").find_elements(By.TAG_NAME, "label")
            for i in range(4):
                development_elements[i].click()
            sleep(0.1)
            
            # entering energy prod
            energy_prod = "50"
            driver.find_element(By.CLASS_NAME, "Section.gcx-forms-section4.unstyled-section").find_element(By.TAG_NAME, "input").send_keys(energy_prod)
            sleep(0.1)
            
            #clicking next
            driver.find_element(By.CLASS_NAME, "ButtonBar.gcx-forms-footer").find_elements(By.TAG_NAME, "button")[1].click()
            sleep(0.2)
            
            # clicking upload area button
            sleep(0.1)
            driver.find_element(By.CLASS_NAME, "RadioGroup.gcx-forms-radioGroup1").find_elements(By.TAG_NAME, "label")[1].click()
            sleep(0.1)
            
            # file upload
            file_input = driver.find_element(By.CLASS_NAME, "FilePicker.gcx-forms-filePicker1").find_element(By.TAG_NAME, "input")
            input_file_target = os.path.join(directory,workingName + ".zip")
            file_input.send_keys(input_file_target)
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CLASS_NAME, "Text.gcx-forms-uploadTextResult")))
            driver.find_element(By.CLASS_NAME, "ButtonBar.gcx-forms-footer").find_elements(By.TAG_NAME, "button")[2].click()
            sleep(0.5)
            
            #starting the report
            driver.find_element(By.CLASS_NAME, "ButtonBar.gcx-forms-footer").find_elements(By.TAG_NAME, "button")[1].click()
            
            # waiting for the report
            pdf_download = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.CLASS_NAME, "toolbar-item.tool.oe-dynamic.pdf-report")))
            pdf_download.click()
            sleep(2) # let it download
            old_file_path = os.path.join(actual_download_path, "RenewableEnergySitingAssessmentProjectInformation_CustomArea.pdf")
            new_file_path = os.path.join(actual_download_path, f"{workingName}.pdf")
            os.rename(old_file_path, new_file_path)
            shutil.move(new_file_path, download_directory)
            os.rmdir(actual_download_path)
            #driver.close()

threads = []
num_threads_and_indice = 2

for i in range(num_threads_and_indice):
    t = FileProcessingThread(i, num_threads_and_indice)
    t.start()
    threads.append(t)
    
for t in threads:
    t.join()