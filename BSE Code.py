from RPA.Browser.Selenium import Selenium
from time import sleep
import gspread
from oauth2client.service_account import ServiceAccountCredentials

browser_lib = Selenium()

url = "https://www.bseindia.com/stock-share-price/tata-steel-ltd/tatasteel/500470/"
google_sheet_url = "https://docs.google.com/spreadsheets/d/1JwqujTpwi5_N8mW5BMXGj5aLOcD4O46BYROdfoSd7aw/edit#gid=0"  # Replace with your Google Sheet URL


def open_the_website(url):
    browser_lib.open_available_browser(url)


def find_closing_price():
    sleep(5)
    closing_price_text = browser_lib.get_text("//strong[@id='idcrval']")
    print("Closing price:", closing_price_text)
    return closing_price_text


def find_date_and_time():
    sleep(5)
    date_time_element = browser_lib.find_element("//div[@class='col-lg-12 ng-binding']")

    if date_time_element:
        date_time_text = date_time_element.text
        date_and_time_parts = date_time_text.split('|')
        date = date_and_time_parts[0].strip()
        time = date_and_time_parts[1].strip()
        print("Date:", date)
        print("Time:", time)
        return date, time
    else:
        print("Error: Could not find date and time elements on the page.")
        return None, None

def find_previous_closing_price():
    sleep(5)
    previous_closing_price_text = browser_lib.get_text("//td[@class='textvalue ng-binding']")
    print("Previous Closing price:", previous_closing_price_text)
    return previous_closing_price_text


def open_the_google_sheet(google_sheet_url):
    browser_lib.open_available_browser(google_sheet_url)

def access_google_sheet(google_sheet_url, closing_price, date, time, previous_closing_price):
    
    credentials = ServiceAccountCredentials.from_json_keyfile_name('C:/Users/Rohan  Dapurkar/Desktop/Projects/BSE Data Scraping/BSE Data Scraping/bse-data-408009-1f61edc80621.json', ['https://spreadsheets.google.com/feeds'])
    gc = gspread.authorize(credentials) #gc = google credentials
    sheet = gc.open_by_url(google_sheet_url)
    worksheet = sheet.get_worksheet(0)

   
    last_row = len(worksheet.col_values(1)) + 1 

    
    
    worksheet.update_acell(f'A{last_row}', closing_price)
    worksheet.update_acell(f'B{last_row}', date)
    worksheet.update_acell(f'C{last_row}', time)
    worksheet.update_acell(f'D{last_row}', previous_closing_price)


open_the_website(url)
closing_price = find_closing_price()
date, time = find_date_and_time()
previous_closing_price = find_previous_closing_price()
open_the_google_sheet(google_sheet_url)
access_google_sheet(google_sheet_url, closing_price, date, time, previous_closing_price)
