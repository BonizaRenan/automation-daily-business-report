from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time

driver = webdriver.Chrome()

driver.get("https://renanbonizaprofile.netlify.app/")

time.sleep(3)  


achievement_tab = driver.find_element(By.XPATH, "//a[contains(text(), 'Experience')]")
achievement_tab.click()

# wait to see result
time.sleep(5)

driver.quit()