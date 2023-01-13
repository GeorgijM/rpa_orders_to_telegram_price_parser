from aiogram import Bot, Dispatcher, types
from aiogram.types import KeyboardButton, ReplyKeyboardMarkup
from aiogram.utils import executor
import logging
import time
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
import price_parser



# Включаем логирование, чтобы не пропустить важные сообщения
logging.basicConfig(level=logging.INFO)
# Объект бота
bot = Bot(token="*******")
# Диспетчер - нужен чтобы реагировать на определенные события ( @dp.message_handler )
dp = Dispatcher(bot)

# настройка скрытого режима браузера для Selenium
chrome_options = Options()
chrome_options.add_argument("--headless")
# chrome_options.add_argument("--disable-notifications")

# название кнопки-клавиатуры в чат-боте
button_price = KeyboardButton('Цены профильной трубы')
# инициализируем появление клавиатуры в чате, resize_keyboard - форматирование размера,
# one_time_keyboard - исчезает после нажатия
price_keyboard = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=False)
# добавляем кнопку в клавиатуру
price_keyboard.add(button_price)


# реагируем на команды "/help" и "/start".
@dp.message_handler(commands=["help", "start"])
async def command_help(message: types.Message):
    # ответ бота, reply_markup отображает выбранную клавиатуру
    await message.answer(" Введите любое слово для поиска картинок или нажмите на кнопку меню для "
                         "получения цен на металл.", parse_mode="HTML", reply_markup=price_keyboard)


# реакция на нажатие кнопки
@dp.message_handler(lambda message: message.text == "Цены профильной трубы")
async def certain_message(msg: types.Message):
    # Ответ бота. Импортируем файл price_parser и вызываем функцию price_parser()
    msg_to_answer = price_parser.price_parser()
    await bot.send_message(msg.from_user.id, msg_to_answer)


# реакция на ввод текста (парсим 10 картинок с депозитфото и публикуем в чате)
@dp.message_handler()
async def query_telegram(msg: types.Message):
    await bot.send_message(msg.from_user.id, "Поиск картинок...")
    print(msg.text)
    driver = webdriver.Chrome(options=chrome_options)
    # открываем сайт
    driver.get("https://ru.depositphotos.com/")
    # ждем пока загрузится
    time.sleep(2)
    # находим элемент по атрибуту name="query"
    textarea = driver.find_element(By.NAME, "query")
    # Вводим текст в найденное выше поле. Используем "\n" для отправки нажатия клавиши "ENTER".
    search_query = msg.text
    textarea.send_keys(f"{search_query}\n")
    time.sleep(2)
    # Находим элемент по css-селектору
    web_elements_with_image = driver.find_elements(By.CSS_SELECTOR, "img")
    time.sleep(1)
    print(web_elements_with_image)
    urls_img = []

    # ошбка при .get_attribute with list (web_elements_with_image.get_attribute), поэтому
    #  используем .get_attribute для каждого отдельного элемента
    for web_element in web_elements_with_image[:10]:
        url_img = web_element.get_attribute('src')
        urls_img.append(url_img)
    print(urls_img)
    # отправляем фото как медиа-группу
    media = types.MediaGroup()
    for i in range(len(urls_img)):
        media.attach_photo(photo=urls_img[i])
    await bot.send_media_group(msg.chat.id, media=media)
    driver.quit()
    print("Done")


if __name__ == "__main__":
    executor.start_polling(dp)
