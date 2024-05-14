from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlwt
import time


service = Service(executable_path="chromedriver.exe")
driver = webdriver.Chrome(service=service)


def scrape(no_of_flats, offers_link):
    #  no_of_flats defines how many flats you want to scrape
    #  offers_link is a link with offers from otodom.pl

    def get_price():
        try:
            price_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'strong[data-cy="adPageHeaderPrice"]')))
            price_text = price_element.text
            price_cleaned = ''.join(filter(str.isdigit, price_text))
            return int(price_cleaned)
        except:
            return "Unknown"


    def get_area():
        try:
            area_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-testid="table-value-area"]')))
            area_text = area_element.text
            area_cleaned = area_text[:-3]
            area_cleaned_with_dot = area_cleaned.replace(",", ".")
            return float(area_cleaned_with_dot)
        except:
            return "Unknown"


    def get_rooms():
        try:
            rooms_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-testid="table-value-rooms_num"]')))
            rooms_text = rooms_element.text
            return int(rooms_text)
        except:
            return "Unknown"

    driver.get(offers_link)

    # Accept cookies
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, 'onetrust-accept-btn-handler')))
        cookies_accept_button = driver.find_element(By.ID, 'onetrust-accept-btn-handler')
        cookies_accept_button.click()
    except:
        pass

    prices = []
    areas = []
    no_of_rooms = []
    links = []

    counter = 0
    page = 1

    while True:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'a[data-cy="listing-item-link"]')))
        links_to_open = [link.get_attribute("href") for link in driver.find_elements(By.CSS_SELECTOR, 'a[data-cy="listing-item-link"]')]
        for link in links_to_open:
            counter += 1
            print(link, counter, page)
            driver.execute_script("window.open('{}', '_blank');".format(link))

            driver.switch_to.window(driver.window_handles[-1])

            prices.append(get_price())
            areas.append(get_area())
            no_of_rooms.append(get_rooms())
            links.append(link)

            driver.close()

            driver.switch_to.window(driver.window_handles[0])

            if counter >= no_of_flats:
                break

        if counter >= no_of_flats:
            break

        try:
            next_page_button = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'li[title="Go to next Page"]')))
            next_page_button.click()
            page += 1
            time.sleep(5)  # No idea why, but without it, the program doesn't fetch the links from the next page
        except:
            break

    driver.quit()

    return prices, areas, no_of_rooms, links


def save_as_xls(prices, areas, no_of_rooms, links):

    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Data')

    for i, (price, area, rooms, link) in enumerate(zip(prices, areas, no_of_rooms, links)):
        sheet.write(i, 0, price)
        sheet.write(i, 1, area)
        sheet.write(i, 2, rooms)
        sheet.write(i, 3, link)

    workbook.save('dane.xls')


if __name__ == '__main__':
    prices, areas, no_of_rooms, links = scrape(50, "https://www.otodom.pl/pl/wyniki/sprzedaz/mieszkanie/malopolskie/krakow/krakow/krakow?viewType=listing&limit=24")
    save_as_xls(prices, areas, no_of_rooms, links)
