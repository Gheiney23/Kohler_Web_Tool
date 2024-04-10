import pandas as pd
import pprint as pp
import time
import re
import urllib.request
from selenium import webdriver as wb
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service

sku_list = [
'Add sku list' 
]

# Setting up the webdriver for Selenium
service = Service()
options = wb.ChromeOptions()
options.add_argument('--start-maximized')
options.add_argument("--disable-notifications")
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = wb.Chrome(service=service, options=options)

opener = urllib.request.build_opener()
opener.addheaders = [('User-Agent', 'MyApp/1.0')]
urllib.request.install_opener(opener)

main_dict = {
        'Sku': [],
        'Collection': [],
        'Finish': [], 
        'Img_url1': [], 
        'Img_url2': [], 
        'Img_url3': [], 
        'Img_url4': [], 
        'Img_url5': [], 
        'Img_url6': [], 
        'Img_url7': [], 
        'Bullet1': [],
        'Bullet2': [],
        'Bullet3': [],
        'Bullet4': [],
        'Bullet5': [],
        'Marketing_copy': [],
        'Specs_Sheet': [],
        'Installation_Sheet': [],
        'Skus_Not_Found': []}

def close_banner_1():
    WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "ip-close")))
    driver.find_element(By.XPATH, "//button[@class='ip-close']").click()


def close_banner_2():
    WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.CLASS_NAME, "bx-button")))
    driver.find_element(By.XPATH, "//button[@class='bx-button']").click()


def kohler_tool():       
        
    # Extracting the image src
    try:
        src_1 = driver.find_element(By.XPATH, "//img[@class='print']").get_attribute('src')
        pp.pprint(src_1)
        main_dict['Img_url1'].append(src_1)
        
    except:
        main_dict['Img_url1'].append('NULL')
    
    # Extracting Collection name
    try:
        collection_ele = driver.find_element(By.XPATH, "//span[@class='product-detail-page__title']").text
        collection = re.sub("[^A-Za-z0-9 -\/]", "", collection_ele)
        main_dict['Collection'].append(collection)
        
    except:
        main_dict['Collection'].append('NULL')

    # Extracting Finish
    try:
        finish = driver.find_element(By.XPATH, "//span[@class='product-detail-page__finish_value']").text
        main_dict['Finish'].append(finish)
        
    except:
        main_dict['Finish'].append('NULL')

    # Extracting Marketing Copy
    try:
        marketing_copy = driver.find_element(By.XPATH, "//p[@class='product-detail-page__narrative-description--show-more']").text
        main_dict['Marketing_copy'].append(marketing_copy)
    
    except:
        main_dict['Marketing_copy'].append('NULL')
    
    # Extracting Bullet points
    try:
        driver.execute_script("window.scrollTo(0, 800)")
        time.sleep(1)
        driver.find_element(By.XPATH, "//div[@class='collapsible__header gbh-data-layer']").click()
        time.sleep(1)

        element = driver.find_element(By.XPATH, "//ul[@class='pdp-features-technologies__feature-list']")
        li_elements = element.find_elements(By.TAG_NAME, 'li')

        b_list = []
        for li in li_elements:
            li_text = li.text
            bullet = re.sub("[^A-Za-z0-9 -\/]", "", li_text)
            b_list.append(bullet)
        
        # Loading all bullet points found into the data_dict
        try:
            main_dict['Bullet1'].append(b_list[0])
        except:
            main_dict['Bullet1'].append('NULL')
        
        try:
            main_dict['Bullet2'].append(b_list[1])
        except:
            main_dict['Bullet2'].append('NULL')
        
        try:
            main_dict['Bullet3'].append(b_list[2])
        except:
            main_dict['Bullet3'].append('NULL')
        
        try:
            main_dict['Bullet4'].append(b_list[3])
        except:
            main_dict['Bullet4'].append('NULL')

        try:
            main_dict['Bullet5'].append(b_list[4])
        except:
            main_dict['Bullet5'].append('NULL')

    except:
        main_dict['Bullet1'].append('NULL')
        main_dict['Bullet2'].append('NULL')
        main_dict['Bullet3'].append('NULL')
        main_dict['Bullet4'].append('NULL')
        main_dict['Bullet5'].append('NULL')

def main(sku_list):
    
    for sku in sku_list:
        try:
            path = 'https://www.kohler.com/en'
            driver.get(path)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "search-icon")))
            time.sleep(1)
            driver.find_element(By.CLASS_NAME, 'search-icon').click()
            time.sleep(1)
            driver.find_element(By.CLASS_NAME, 'search-side-panel__search-control').send_keys(sku)
            driver.find_element(By.CLASS_NAME, 'search-side-panel__search-control').send_keys(Keys.RETURN)
            time.sleep(3)
            main_dict['Sku'].append(sku)

            driver.switch_to.parent_frame()
            kohler_tool()
            
        except:
                main_dict['Sku'].append(sku)
                main_dict['Collection'].append('NULL')
                main_dict['Finish'].append('NULL')
                main_dict['Img_url1'].append('NULL')
                main_dict['Img_url2'].append('NULL')
                main_dict['Img_url3'].append('NULL')
                main_dict['Img_url4'].append('NULL')
                main_dict['Img_url5'].append('NULL')
                main_dict['Img_url6'].append('NULL')
                main_dict['Img_url7'].append('NULL')
                main_dict['Bullet1'].append('NULL')
                main_dict['Bullet2'].append('NULL')
                main_dict['Bullet3'].append('NULL')
                main_dict['Bullet4'].append('NULL')
                main_dict['Bullet5'].append('NULL')
                main_dict['Marketing_copy'].append('NULL')
                main_dict['Specs_Sheet'].append('NULL')
                main_dict['Installation_Sheet'].append('NULL')
                main_dict['Skus_Not_Found'].append(sku)

    # quitting the driver and manipulation the dictionary into a dataframe
    driver.quit()

    df = pd.DataFrame.from_dict(main_dict,orient='index')
    df = df.transpose()
    # # df['Img_url'].fillna('NULL', inplace=True)

    # Writing the dataframe to an excel worksheet
    df.to_excel('Kohler_Data.xlsx', sheet_name='Kohler_Data')

    print('Run Complete!')

# Run the program
if __name__ == "__main__":
    main(sku_list)