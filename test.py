import time
import keyboard
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import chromedriver_autoinstaller



service = Service()
profile_path = r''  #private information
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument(f'--user-data-dir={profile_path}')
driver = webdriver.Chrome(service=service, options=chrome_options)
#login ortus
login=""#private information
password=""#private information
driver.get('https://id2.rtu.lv/')
login_f=driver.find_element(By.NAME,"IDToken1")
login_f.click()
login_f.clear()
login_f.send_keys(login)
password_f=driver.find_element(By.ID,"IDToken2")
password_f.send_keys(password)
login_f=driver.find_element(By.NAME,"Login.Submit")
login_f.click()
#MAIL block
driver.get('https://mail.google.com/mail')
#parent=driver.getWindowHandle()




time.sleep(1)

search_input = driver.find_element(By.NAME, "q")

search_input.click()
search_input.send_keys("pievienojis atsauksmi")


search_input.send_keys(webdriver.Keys.RETURN)

time.sleep(3)
email_open= driver.find_element(By.CLASS_NAME, 'av')
email_open.click()
time.sleep(1)
link_open= driver.find_element(By.LINK_TEXT,'uzdevuma iesniegumam')
link_open.click()
time.sleep(5)
keyboard.press(['ctrl','p'])
keyboard.release('ctrl')
keyboard.release('p')
time.sleep(3) 
keyboard.press_and_release('enter')
time.sleep(3)
keyboard.press_and_release('enter')
time.sleep(10)


driver.quit()