from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
from openpyxl import load_workbook

chrome_options = Options()
# Раскоментить если нужно скрыть окно браузера
# chrome_options.add_argument("--headless")
# Раскоментить если на сайте будут всплывающие окна
# chrome_options.add_argument("--disable-notifications")


# Инициализируем переменные для значений цены за метр и тону. Это нужно потому, что
# далее в цикле с помощью команды globals() мы будем присваивать название переменной
# согласно названию ключа из словаря {links}
price_40x20x015_per_meter = ""
price_40x20x015_per_ton = ""
price_40x20x020_per_ton = ""
price_40x20x020_per_meter = ""
price_60x40x015_per_ton = ""
price_60x40x015_per_meter = ""
price_60x40x020_per_ton = ""
price_60x40x020_per_meter = ""
price_60x60x020_per_ton = ""
price_60x60x020_per_meter = ""
price_80x80x030_per_ton = ""
price_80x80x030_per_meter = ""
price_100x100x030_per_ton = ""
price_100x100x030_per_meter = ""


def price_parser():
    # Инициализируем вебдрайвер и открываем окно браузера(Chrome)
    driver = webdriver.Chrome(options=chrome_options)
    # Вносим данные в словарь. Ключи - виды труб - будем подставлять в название переменной в цикле прохода по ссылкам.
    # Значения - ссылки на страницы, откуда будем вытягивать цены для определенного вида трубы
    links = {
        "price_40x20x015": "https://aksvil.by/truby-profilnyye/truba-profilnaya-40x20x1-5-mm.html",
        "price_40x20x020": "https://aksvil.by/truby-profilnyye/truba-profilnaya-40x20x2-mm.html",
        "price_60x40x015": "https://aksvil.by/truby-profilnyye/truba-profilnaya-60-40-1-5-mm.html",
        "price_60x40x020": "https://aksvil.by/truby-profilnyye/truba-profilnaya-60x40x2-mm.html",
        "price_60x60x020": "https://aksvil.by/truby-profilnyye/truba-profilnaya-60x60x2mm.html",
        "price_80x80x030": "https://aksvil.by/truby-profilnyye/truba-profilnaya-80x80x3-mm.html",
        "price_100x100x030": "https://aksvil.by/truby-profilnyye/truba-profilnaya-100x100x3-mm.html",
    }

    for i in range(len(links)):
        # Получаем i-ую ссылку из словаря
        link_for_search = list(links.values())[i]
        # Открываем ссылку в браузере
        driver.get(link_for_search)
        # Находим элементы HTML страницы у которых css-селектор "div.price"
        price_all_from_page = driver.find_elements(By.CSS_SELECTOR, "div.price")
        # Берем первый элемент из всех найденых для цены за тону
        price_per_ton = price_all_from_page[0].text
        # Берем второй элемент из всех найденых для цены за метр
        price_per_meter = price_all_from_page[1].text
        # Составляем имя переменной для цены за тону на определенной странице.
        # Для этого приводим ключи links.keys() из словаря к списку, затем берем i-ое значение (по циклу)
        # добавляем к нему "_per_ton" и преобразуем в строку.
        name_for_variable_per_ton = str(list(links.keys())[i] + "_per_ton")
        # Аналогично строке выше, но для цены за метр
        name_for_variable_per_meter = str(list(links.keys())[i] + "_per_meter")
        # Сохраняем значение переменных с выбранными выше названиями за рамками цикла
        globals()[name_for_variable_per_ton] = price_per_ton
        globals()[name_for_variable_per_meter] = price_per_meter
    # выводим цены в консоль
    print(f'{price_40x20x015_per_meter=}')
    print(f'{price_40x20x015_per_ton=}')
    print(f'{price_40x20x020_per_meter=}')
    print(f'{price_40x20x020_per_ton=}')
    print(f'{price_60x40x015_per_meter=}')
    print(f'{price_60x40x015_per_ton=}')
    print(f'{price_60x40x020_per_meter=}')
    print(f'{price_60x40x020_per_ton=}')
    print(f'{price_60x60x020_per_meter=}')
    print(f'{price_60x60x020_per_ton=}')
    print(f'{price_80x80x030_per_meter=}')
    print(f'{price_80x80x030_per_ton=}')
    print(f'{price_100x100x030_per_meter=}')
    print(f'{price_100x100x030_per_ton=}')
    # Закрываем окно браузера
    driver.quit()

    # Строка для возврата функцией (используем в телеграм боте)
    price_list = f'40х20х1,5: {price_40x20x015_per_meter}; {price_40x20x015_per_ton}\n' \
                 f'40х20х2,0: {price_40x20x020_per_meter}; {price_40x20x020_per_ton}\n' \
                 f'60х40х1,5: {price_60x40x015_per_meter}; {price_60x40x015_per_ton}\n' \
                 f'60х40х2,0: {price_60x40x020_per_meter}; {price_60x40x020_per_ton}\n' \
                 f'60х60х2,0: {price_60x60x020_per_meter}; {price_60x60x020_per_ton}\n' \
                 f'80х80х3,0: {price_80x80x030_per_meter}; {price_80x80x030_per_ton}\n' \
                 f'100х100х3,0: {price_100x100x030_per_meter}; {price_100x100x030_per_ton}\n'
    return price_list


# Функция для приведения значений цены из string во float
def price_to_float(price_string):
    # Находим в строке индекс подстроки " руб."
    slash_index = price_string.find('руб')
    # Берем срез строки до подстроки " руб"
    price_slice = price_string[:slash_index]
    # Удаляем пробелы в значении цены, чтобы привести во float
    price_slice = price_slice.replace(' ', '')
    # Из string во float
    price_to_float = float(price_slice)
    return price_to_float


# Вносим значения цен с сайта в эксель таблицу.
# Не через цикл, чтобы быстрее настроить в случае изменения эксель-файла
def set_excel_prices():
    wb = load_workbook('смета_цены.xlsx')
    sheet = wb['материалы цена']
    sheet['C2'] = price_to_float(price_40x20x015_per_meter)
    sheet['C3'] = price_to_float(price_40x20x020_per_meter)
    sheet['C4'] = price_to_float(price_60x40x015_per_meter)
    sheet['C5'] = price_to_float(price_60x40x020_per_meter)
    sheet['C6'] = price_to_float(price_60x60x020_per_meter)
    sheet['C7'] = price_to_float(price_80x80x030_per_meter)
    sheet['C8'] = price_to_float(price_100x100x030_per_meter)
    # print(sheet['C8'].value)
    wb.save('смета_цены.xlsx')


if __name__ == "__main__":
    price_parser()
    set_excel_prices()
