
# from selenium import webdriver 
# from selenium.webdriver.support.ui import WebDriverWait 
# from selenium.webdriver.support import expected_conditions as EC 
# from selenium.webdriver.common.keys import Keys 
# from selenium.webdriver.common.by import By 
# import time 
  
# # Replace below path with the absolute path 
# # to chromedriver in your computer 
# driver = webdriver.Chrome(r'C:\Users\ucc14\Downloads\chromedriver_win32\chromedriver.exe') 
  
# driver.get("https://web.whatsapp.com/") 
# wait = WebDriverWait(driver, 600) 
  
# # Replace 'Friend's Name' with the name of your friend  
# # or the name of a group  
# target = '"Gagan"'
  
# # Replace the below string with your own message 
# string = "Message sent using Python!!!"
  
# x_arg = '//span[contains(@title,' + target + ')]'
# group_title = wait.until(EC.presence_of_element_located(( 
#     By.XPATH, x_arg))) 
# group_title.click() 
# inp_xpath = '//div[@class="input"][@dir="auto"][@data-tab="1"]'
# input_box = wait.until(EC.presence_of_element_located(( 
#     By.XPATH, inp_xpath))) 
# for i in range(100): 
#     input_box.send_keys(string + Keys.ENTER) 
#     time.sleep(1)


# from selenium import webdriver
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.by import By
# import time
# import sys
# # Replace below path with the absolute path of the \
# #chromedriver in your computer
# driver = webdriver.Chrome(r'C:\Users\ucc14\Downloads\chromedriver_win32\chromedriver.exe')

# driver.get("https://web.whatsapp.com/")
# # time.sleep()
# wait = WebDriverWait(driver, 600)
# # Replace 'My Bsnl' with the name of your friend or group name
# target = '"A Mammi Ji"'

# # Replace the below string with your own message
# string = sys.argv[1]
# # string = "Hi this msg is from Python"

# x_arg = '//span[contains(@title,' + target + ')]'
# group_title = wait.until(EC.presence_of_element_located((
# By.XPATH, x_arg)))
# print (group_title)
# print ("Wait for few seconds")
# group_title.click()
# message = driver.find_elements_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')[0]

# message.send_keys(string)
# sendbutton = driver.find_elements_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button')[0]
# sendbutton.click()
# driver.close()


