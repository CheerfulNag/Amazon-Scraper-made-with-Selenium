#Importing_modules
from selenium import webdriver
from bs4 import BeautifulSoup
import time
import xlsxwriter
import pandas as pd
import requests
import concurrent.futures 
import os
import shutil
import warnings


#------------------------------------------------------
#Prepatation_modules
def country_and_searchterm(search_term,country_code):
    #Use country selection function and search site for search_term
    skip()
    time.sleep(1)
    country_selecting(country_code)
    driver.refresh()
    search_box = driver.find_element_by_xpath('//*[@id="twotabsearchtextbox"]')
    search_box.send_keys(search_term)
    search_button = driver.find_element_by_xpath('//*[@id="nav-search-submit-button"]')
    search_button.click()
    time.sleep(3)

def country_selecting(country_code):
    #Change deliver country
    country_button = driver.find_elements_by_xpath('//*[@id="glow-ingress-line2"]')[0]
    country_button.click()
    time.sleep(1)

    country_button2 = driver.find_elements_by_xpath('/html/body/div[3]/div/div/div[1]/div/div[2]/div[3]/div[4]/span/span/span/span')[0]
    country_button2.click()
    time.sleep(1)

    country_button3 = driver.find_elements_by_xpath(f'//*[@id="GLUXCountryList_{country_code}"]')[0]
    country_button3.click()
    
def skip():
    #Skips pop-up's when changes country
    try:      
        if driver.find_element_by_xpath('/html/body/div[1]/header/div/div[4]/div[1]/div/div/div[3]/span[1]/span/input'):
            button = driver.find_element_by_xpath('/html/body/div[1]/header/div/div[4]/div[1]/div/div/div[3]/span[1]/span/input')
            button.click()
    except:
        pass
    try:      
        if driver.find_element_by_xpath('/html/body/div[1]/header/div/div[3]/div[13]/div[2]/div[4]/span[1]/span/input'):
            button = driver.find_element_by_xpath('/html/body/div[1]/header/div/div[3]/div[13]/div[2]/div[4]/span[1]/span/input')
            button.click()
    except:
        pass
#------------------------------------------------------


#------------------------------------------------------
#Extraction_part
def extraction_from_page():
    #Extract products data from page
    soup = BeautifulSoup(driver.page_source, "html.parser")
    results = soup.find_all('div', {'data-component-type':"s-search-result"})

    for item in results:
        if item.find('div', {'data-component-type':"sp-sponsored-result"}) == None:
            try:
                price = item.find('span', {'class':'a-offscreen'}).text
                price = float(price.replace(",","").replace('$',''))
            except:
                price = 'Unkown'
            try:
                rating = item.find('span',{'class':'a-icon-alt'}).text
                reviews = int((item.find('span',{'class':'a-size-base'}).text).replace(',',''))
            except:
                rating = 'Unkown'
                reviews = 'Unkown'
            try:
                label = item.find('a',{'class':'a-link-normal a-text-normal'})
                name = label.text.strip()
                url = 'https://www.amazon.com/'+label.get('href') 
            except:
                label = item.find('a',{'class':'a-link-normal s-underline-text s-underline-link-text a-text-normal'})
                name = label.text.strip()
                url = 'https://www.amazon.com/'+label.get('href')
            record = (name,price,rating,reviews,url)
            records.append(record)
            link = (item.find('img',{'class':'s-image'})).get('src')
            links.append(link)

def next_page(render_time):
    #Pressing next_page button
    try:
        try:
            next_page_button = driver.find_elements_by_css_selector('a[class="s-pagination-item s-pagination-next s-pagination-button s-pagination-separator"]')[0]
            next_page_button.click()
            time.sleep(render_time)
            return 1    
        except:
            next_page_button = driver.find_elements_by_css_selector('li[class="a-last"')[0]
            next_page_button.click()
            time.sleep(render_time)
            return 1
    except:
        time.sleep(render_time)
        return 0

def loader(url):
    with open(str(links.index(url))+ ".jpg",'wb') as f:
        im = requests.get(url)
        f.write(im.content)
        time.sleep(0.5)
        f.close()
    
def img_loader():   
    with concurrent.futures.ThreadPoolExecutor() as executor:
        executor.map(loader,links)
#------------------------------------------------------


#------------------------------------------------------
#Data_saving_part
def xlsx_saver(search_term,country_code):
    #Saves data into excel
    name = (search_term.replace(" ","_")+f"_{countries_codes_dict[country_code]}_results.xlsx")
    workbook = xlsxwriter.Workbook(name) 

    #Formats creation
    align_format = workbook.add_format()
    align_format.set_align('fill')
    align_format.set_border()
    currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    currency_format.set_border()
    headings_format = workbook.add_format()
    headings_format.set_bg_color('#C0C0C0')
    headings_format.set_border()
    cell_format = workbook.add_format()
    cell_format.set_border()
    n = 2

    #Writing
    worksheet = workbook.add_worksheet()
    worksheet.set_default_row(90)
    worksheet.write('A1', "Name",headings_format)  
    worksheet.set_column("A:A",50)
    worksheet.write('B1', "Price",headings_format)
    worksheet.write('C1', "Rating",headings_format)
    worksheet.set_column("C:C",10)
    worksheet.write('D1', "Reviews",headings_format)
    worksheet.write('E1', "URL",headings_format)
    worksheet.set_column("E:E",30)
    worksheet.write('F1', "Image",headings_format)
    worksheet.set_column("F:F",30)


    for record in records:
        worksheet.write(f'A{n}', record[0],cell_format)
        worksheet.write(f'B{n}', record[1],currency_format)
        worksheet.write(f'C{n}', record[2],cell_format)
        worksheet.write(f'D{n}', record[3],cell_format)
        worksheet.write(f'E{n}', record[4],align_format)
        warnings.filterwarnings("ignore")
        worksheet.insert_image(f'F{n}',
                            f'{search_term}_images/{n-2}.jpg', 
                           {'x_scale': 0.5, 'y_scale': 0.5, 
                            'x_offset': 5, 'y_offset': 5,
                            'positioning': 1})


        n+=1
    time.sleep(3)
    workbook.close()

    print("EXCEL FILE COMPLETE")

    #Results check
    df = pd.read_excel(name)
    pd.set_option('max_rows',5000)
    print(df)
#------------------------------------------------------


#------------------------------------------------------
#Main_functions 
def main_extraction(search_term,country_code,render_time):
    try:
        driver.get('https://www.amazon.com/')
        time.sleep(1)
        country_and_searchterm(search_term,country_code)
        while True:
            extraction_from_page()
            w = next_page(render_time)
            if w == False:
                break
        driver.close()
        print("Scraping is done")
    except:
        print("Scraping went wrong, please try again")

def main(search_term,country_code,del_or_not = None,render_time = 1.5):
    start = time.time()

    global links 
    global records
    global options
    global driver
    global countries_codes_dict
    countries_codes_dict = {
    "0":"Australia",
    "1":"Canada",
    "2":"China",
    "3":"Japan",
    "4":"Mexico",
    "5":"Singapore",
    "6":"UK",
    }
    links = []
    records = []
    options = webdriver.FirefoxOptions()
    #options.headless = True
    driver = webdriver.Firefox(options = options,executable_path='geckodriver.exe')

    main_extraction(search_term,country_code,int(render_time))

    main_dir = os.getcwd()
    try:
        os.mkdir(search_term+"_images")
        print("Temporary folder created")
    except:
        pass
    os.chdir(search_term+"_images")
    img_loader()
    os.chdir(main_dir)

    xlsx_saver(search_term,str(country_code))
    if del_or_not == True:
        pass
    else:
        shutil.rmtree(search_term+"_images")

    end = time.time()
    total = time.gmtime(end-start)
    total_time = time.strftime("%M:%S",total) 
    print("Total: ",total_time)
#------------------------------------------------------




main("gaming laptop",0)
main("gaming monitor",1)
main("gaming keyboard",6)



#################INSTRUCTIONS#################
#Input yout search_term first, then country code you need, then if you want to 
#save images outside the excel file, input "True".
#If your internet is slow, you can change render_time to a bigger value, or you'll notice that not all results are found(default value equal 1.5).
#P.S Also it can be just one time lag. + about 1-5% of all images(always the same for one search_term) are not loaded for unknown reasons

#Should look like: main('my_search_term',3,True,render_time = 3)
#You can make several runs in a row


#Country_codes:
# Australia = 0
# Canada = 1
# China = 2
# Japan = 3
# Mexico = 4
# Singapore = 5 
# United Kingdom = 6





