import openpyxl
import requests
from bs4 import BeautifulSoup

def scrape_amazon_products(url, num_pages=1):
    all_products_data = []

    for page in range(1, num_pages + 1):
        page_url = f"{url}&page={page}"
        response = requests.get(page_url)

        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            product_items = soup.find_all('div', {'class': 's-card-container'})

            for item in product_items:
                name_element = item.find('span', {'class': 'a-size-medium a-color-base a-text-normal'})
                name = name_element.text.strip() if name_element else ""

                price_element = item.find('span', {'class': 'a-price-whole'})
                price = price_element.text.strip() if price_element else ""

                rating_element = item.find('span', {'class': 'a-icon-alt'})
                rating = rating_element.text.strip().split()[0] if rating_element else ""

                num_reviews_element = item.find('span', {'class': 'a-size-base s-underline-text'})
                num_reviews = num_reviews_element.text.strip() if num_reviews_element else ""

                all_products_data.append([name, price, rating, num_reviews])

    return all_products_data

def write_to_excel(data, file_name):
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    sheet.append(["Name", "Price", "Rating", "Number of Reviews"])
    
    for product in data:
        sheet.append(product)

    workbook.save(file_name)

if __name__ == "__main__":
    search_url = "https://www.amazon.in/s?k=bags&crid=2M096C61O4MLT&qid=1653308124&sprefix=ba%2Caps%2C283&ref=sr_pg_1"
    num_pages_to_scrape = 20
    scraped_data = scrape_amazon_products(search_url, num_pages_to_scrape)

    if scraped_data:
        write_to_excel(scraped_data, "amazon_products_data.xlsx")
        print("Data saved to amazon_products_data.xlsx")
    else:
        print("No data extracted.")
