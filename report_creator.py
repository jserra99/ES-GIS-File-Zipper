''' Created by Joseph Serra, student in EECS @Oregon State University. 
Source: https://github.com/jserra99/ES-GIS-File-Zipper '''

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import os
import threading
import shutil

download_directory = str(input("Please copy/paste the export directory path: "))

url = 'https://tools.oregonexplorer.info/OE_HtmlViewer/Index.html?viewer=renewable'

directory = str(input("Please copy/paste the source directory path containing the zip files: "))

temp = []
for root, directories, files, in os.walk(directory):
        for file in files:
            temp.append(file)

completed_list = []
for root, directories, files, in os.walk(download_directory):
        for file in files:
            completed_list.append(file)


zip_list = []
for fileName in temp:
    if str(fileName).removesuffix(".zip") + ".pdf" not in completed_list:
        zip_list.append(fileName)
zip_list_len = len(zip_list)

class FileProcessingThread(threading.Thread):
    def __init__(self, currentIndex: int, jump: int):
        threading.Thread.__init__(self)
        self.currentIndex = currentIndex
        self.jump = jump
        
    
    def run(self):
        
        while(self.currentIndex < zip_list_len):
            options = Options()
            options.add_experimental_option("detach", True)
            workingName = str(zip_list[self.currentIndex]).removesuffix('.zip')
            chromeDriverManager = ChromeDriverManager()
            actual_download_path = os.path.join(download_directory,workingName)
            # print(f"Download Path index: {self.currentIndex}, Path: {actual_download_path}")
            options.add_experimental_option('prefs', {
            "download.default_directory": actual_download_path, #Change default directory for downloads
            "download.prompt_for_download": False, #To auto download the file
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True #It will not show PDF directly in chrome
            })
            os.makedirs(actual_download_path, exist_ok=True)
            driver = webdriver.Chrome(service=Service(chromeDriverManager.install()), 
                                options=options)
            file_processing_done = False
            while not file_processing_done:
                try:
                    driver.get(url)

                    # clicking splash screen button, webdriver wait doesn't work here because of the initial loading screen, waiting is the only way around it
                    sleep(20)
                    driver.find_elements(By.CLASS_NAME, "oeFormActionsRow")[1].find_elements(By.TAG_NAME, "button")[1].click()
                    sleep(2)

                    # clicking disclaimer next button
                    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, "ButtonBar.gcx-forms-footer")))
                    driver.find_element(By.CLASS_NAME, "ButtonBar.gcx-forms-footer").find_elements(By.TAG_NAME, "Button")[1].click()
                    sleep(1)

                    # filling out the site name & other info
                    # site name
                    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "gcx-forms-checkGroupDevTypesOnShore-18")))
                    driver.find_element(By.CLASS_NAME, "Body").find_element(By.TAG_NAME, "input").send_keys(workingName)
                    sleep(0.25)

                    # doing development checkboxes
                    development_elements = driver.find_element(By.ID, "gcx-forms-checkGroupDevTypesOnShore-18").find_element(By.CLASS_NAME, "items").find_elements(By.TAG_NAME, "label")
                    for i in range(4):
                        development_elements[i].click()
                    sleep(0.25)
                    
                    # entering energy prod
                    energy_prod = "50"
                    driver.find_element(By.CLASS_NAME, "Section.gcx-forms-section4.unstyled-section").find_element(By.TAG_NAME, "input").send_keys(energy_prod)
                    sleep(0.25)
                    
                    #clicking next
                    driver.find_element(By.CLASS_NAME, "ButtonBar.gcx-forms-footer").find_elements(By.TAG_NAME, "button")[1].click()
                    sleep(0.5)
                    
                    # clicking upload area button
                    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, "RadioGroup.gcx-forms-radioGroup1")))
                    driver.find_element(By.CLASS_NAME, "RadioGroup.gcx-forms-radioGroup1").find_elements(By.TAG_NAME, "label")[1].click()
                    sleep(0.5)
                    
                    # file upload
                    file_input = driver.find_element(By.CLASS_NAME, "FilePicker.gcx-forms-filePicker1").find_element(By.TAG_NAME, "input")
                    input_file_target = os.path.join(directory,workingName + ".zip")
                    file_input.send_keys(input_file_target)
                    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, "Text.gcx-forms-uploadTextResult")))
                    driver.find_element(By.CLASS_NAME, "ButtonBar.gcx-forms-footer").find_elements(By.TAG_NAME, "button")[2].click()
                    sleep(0.5)
                    
                    #starting the report
                    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.CLASS_NAME, "ButtonBar.gcx-forms-footer")))
                    driver.find_element(By.CLASS_NAME, "ButtonBar.gcx-forms-footer").find_elements(By.TAG_NAME, "button")[1].click()
                    
                    # waiting for the report
                    pdf_download = WebDriverWait(driver, 360).until(EC.presence_of_element_located((By.CLASS_NAME, "toolbar-item.tool.oe-dynamic.pdf-report")))
                    pdf_download.click()
                    
                    old_file_path = os.path.join(actual_download_path, "RenewableEnergySitingAssessmentProjectInformation_CustomArea.pdf")
                    while not os.path.exists(old_file_path):
                        sleep(1)
                    new_file_path = os.path.join(actual_download_path, f"{workingName}.pdf")
                    os.rename(old_file_path, new_file_path)
                    shutil.move(new_file_path, download_directory)
                    os.rmdir(actual_download_path)
                    file_processing_done = True
                    driver.close()
                    self.currentIndex += self.jump
                except:
                    pass #driver.close()
                    
class FileProgressThread(threading.Thread):
    def __init__(self, completed_list_len, og_zip_list_len):
        threading.Thread.__init__(self)
        print(f"Percent Complete: {round((100 * completed_list_len / og_zip_list_len), 2)}%")
        self.og_zip_list_len = og_zip_list_len
        
    def run(self):
        percentComplete = 0

        while round(percentComplete, 2) != 100:
            sleep(30)
            completed_list = []
            for root, directories, files, in os.walk(download_directory):
                    for file in files:
                        completed_list.append(file)
            percentComplete = round((100 * len(completed_list) / self.og_zip_list_len), 2)
            print(f"Percent Complete: {percentComplete}%")

threads = []
num_threads_and_indice = int(input("How many threads would you like to run?: "))

for i in range(num_threads_and_indice):
    t = FileProcessingThread(i, num_threads_and_indice)
    t.start()
    threads.append(t)
t_main = FileProgressThread(len(completed_list), len(temp))
t_main.start()
threads.append(t_main)

for t in threads:
    t: threading.Thread
    t.join()
    