import openpyxl
import os

from playwright.sync_api import sync_playwright


def load_products(filename: str) -> list:
    wb = openpyxl.load_workbook(f'{filename}.xlsx')
    sheet = wb.active['A:A']
    products = []

    for product in sheet:
        products.append(product.value)

    return products


def wright_price(price_avito: float, price_yandex: float, name: str) -> None:
    if not os.path.exists('output.csv'):
        with open('output.csv', 'w+', encoding='utf-8') as file:
            file.write('Название, ЯндексМаркет, Авито, -30%, -45%, -70%\n')

    with open('output.csv', 'a', encoding='utf-8') as file:
        file.write(f'{name}, {price_yandex},'
                   f' {price_avito},'
                   f' {price_yandex * 0.7},'
                   f' {price_yandex * 0.55},'
                   f' {price_yandex * 0.3}\n')


def pars_market(products: list) -> None:
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        for product in products:
            # парсинг авито
            product.replace(' ', '+')

            page.goto(f"https://www.avito.ru/all?f=ASgCAgECAUX"
                      f"GmgwReyJmcm9tIjoxLCJ0byI6MH0&q={product}")

            for i in range(24):
                page.keyboard.press('PageDown')

            box = page.query_selector('.items-items-kAJAg')
            items = box.query_selector_all('[data-marker=item]')
            average = []
            for item in items:
                price = item.query_selector('.iva-item-priceStep-uq2CQ').inner_text()

                clear_price = ''
                price = price.split('—')[0]
                for i in price:
                    if i in ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0']:
                        clear_price += i

                average.append(int(clear_price))

            if len(average) != 0:
                average_avito = int(sum(average) / len(average))
            else:
                average_avito = 0
            print(f'{product}: \n   Авито -> {average_avito}')
            # парсинг яндекс маркет
            page.goto("https://market.yandex.ru")
            page.get_by_placeholder("Искать товары").click()
            page.get_by_placeholder("Искать товары").fill(product.replace('+', ' '))
            page.get_by_role("button", name="Найти").click()
            page.wait_for_timeout(3000)

            for i in range(12):
                page.keyboard.press('PageDown')
                page.wait_for_timeout(100)

            average = []
            items = page.query_selector_all('.nXZ_7')
            for item in items:
                price = item.query_selector('._1stjo')
                if price is not None:
                    price = price.inner_text().replace('\u2009', ' ').replace('\u202f', ' ')
                    clear_price = ''
                    for i in price:
                        if i in ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0']:
                            clear_price += i
                    average.append(int(clear_price))

            if len(average) != 0:
                averager_yandex = int(sum(average) / len(average))
            else:
                averager_yandex = 0

            print(f'   ЯндексМаркет -> {averager_yandex}')
            wright_price(average_avito, averager_yandex, product)
