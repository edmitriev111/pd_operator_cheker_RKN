import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import xml.etree.ElementTree as ET
from xml.dom import minidom
import pandas as pd
import os
import sys
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

def read_inn_list(filename):
    with open(filename, 'r') as file:
        return [line.strip() for line in file if line.strip()]

def verify_chrome_installation():
    import winreg
    import os
    
    try:
        # Проверка версии Chrome
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
        version = winreg.QueryValueEx(key, "version")[0]
        
        # Проверка пути установки
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        if not os.path.exists(chrome_path):
            chrome_path = r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
            
        return {
            'version': version,
            'path': chrome_path if os.path.exists(chrome_path) else None,
            'is_valid': os.path.exists(chrome_path)
        }
    except Exception as e:
        return {'error': str(e)}

def setup_driver():
    try:
        print("Начало инициализации ChromeDriver...")
        
        # Настройка опций Chrome
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--window-size=1920,1080')
        
        # Используем ChromeDriverManager без явного указания версии
        driver_path = ChromeDriverManager().install()
        
        # Проверяем и корректируем путь к chromedriver.exe
        if not driver_path.endswith('chromedriver.exe'):
            chrome_dir = os.path.dirname(driver_path)
            for file in os.listdir(chrome_dir):
                if file.endswith('chromedriver.exe'):
                    driver_path = os.path.join(chrome_dir, file)
                    break
        
        print(f"Путь к ChromeDriver: {driver_path}")
        
        # Создаем службу с проверенным путем
        service = Service(executable_path=driver_path)
        
        # Инициализация драйвера
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("ChromeDriver успешно инициализирован")
        return driver
            
    except Exception as e:
        print(f"Критическая ошибка при инициализации ChromeDriver: {str(e)}")
        print("\nДля исправления выполните команды:")
        print("1. Удалите папку:")
        print("   rmdir /s /q %USERPROFILE%\\.wdm")
        print("2. Выполните установку драйвера:")
        print("   pip install --upgrade webdriver-manager")
        return None

def extract_email(contact_text):
    """Извлекает email из текста контактных данных"""
    if not contact_text:
        return ''
    lines = contact_text.split('\n')
    for line in lines:
        if '@' in line:
            return line.strip()
    return ''

def check_operator_status(driver, inn):
    try:
        url = 'https://pd.rkn.gov.ru/operators-registry/operators-list/'
        print(f"Открываем страницу: {url}")
        driver.get(url)
        
        wait = WebDriverWait(driver, 15)
        
        # Ждем поле ввода ИНН
        print("Ожидаем поле ввода ИНН...")
        inn_input = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input[name='inn']"))
        )
        print("Поле ввода ИНН найдено")
        
        # Вводим ИНН
        inn_input.clear()
        inn_input.send_keys(inn)
        print(f"Введен ИНН: {inn}")
        
        # Отправляем форму
        print("Отправляем форму...")
        form = driver.find_element(By.TAG_NAME, "form")
        form.submit()
        
        # Ждем появления результатов
        print("Ожидаем результаты...")
        try:
            table = wait.until(
                EC.presence_of_element_located((By.ID, "ResList1"))
            )
            
            # Получаем строки с данными
            rows = table.find_elements(By.XPATH, ".//tbody/tr")
            if not rows:
                print(f"Для ИНН {inn} нет данных в таблице")
                return None
            
            # Сохраняем данные из таблицы в переменные
            cells = None
            detail_url = None
            
            # Получаем данные из первой строки и сохраняем их
            cells = rows[0].find_elements(By.TAG_NAME, "td")
            if len(cells) < 6:
                print(f"Неверное количество столбцов в таблице: {len(cells)}")
                return None
                
            # Получаем ссылку на детали до перехода на новую страницу
            detail_link = cells[1].find_element(By.TAG_NAME, "a")
            detail_url = detail_link.get_attribute("href")
            
            # Сохраняем основные данные
            result = {
                'reg_number': cells[0].text.strip(),
                'name_inn': cells[1].text.strip(),
                'operator_type': cells[2].text.strip(),
                'inclusion_basis': cells[3].text.strip(),
                'registration_date': cells[4].text.strip(),
                'processing_start_date': cells[5].text.strip(),
                'url': detail_url
            }
            
            # Переходим на страницу деталей
            print(f"Переход на страницу деталей: {detail_url}")
            driver.get(detail_url)
            time.sleep(2)  # Даем время на загрузку страницы
            
            try:
                # Ищем информацию об ответственном лице с использованием полного XPath
                responsible_xpath = "//td[contains(text(), 'ФИО физического лица или наименование юридического лица, ответственных за организацию обработки персональных данных')]/following-sibling::td"
                responsible_element = wait.until(
                    EC.presence_of_element_located((By.XPATH, responsible_xpath))
                )
                responsible_person = responsible_element.text.strip()
                
                # Ищем контактную информацию
                contacts_xpath = "//td[contains(text(), 'номера их контактных телефонов, почтовые адреса и адреса электронной почты')]/following-sibling::td"
                contacts_element = wait.until(
                    EC.presence_of_element_located((By.XPATH, contacts_xpath))
                )
                contact_details = contacts_element.text.strip()
                email = extract_email(contact_details)
                
                # Добавляем найденную информацию в результат
                result.update({
                    'responsible_person': responsible_person,
                    'contact_details': contact_details,
                    'email': email  # Добавляем email в результат
                })
                
                print(f"Email получен: {email}")
                
            except Exception as e:
                print(f"Ошибка при получении детальной информации: {str(e)}")
                result.update({
                    'responsible_person': 'Не указано',
                    'contact_details': 'Не указано',
                    'email': ''
                })
            
            return result
            
        except Exception as e:
            print(f"Ошибка при поиске данных в таблице: {str(e)}")
            return None
            
    except Exception as e:
        print(f"Ошибка при проверке ИНН {inn}: {str(e)}")
        print("\nHTML страницы:")
        print(driver.page_source)
        return None

def create_xml_report(results):
    root = ET.Element('Операторы')
    
    for result in results:
        operator = ET.SubElement(root, 'Оператор')
        
        if result['data']:
            # Разделяем название и ИНН
            name_inn = result['data']['name_inn']
            name = name_inn.split('ИНН:')[0].strip()
            inn = name_inn.split('ИНН:')[1].strip() if 'ИНН:' in name_inn else ''
            
            # Основная информация
            ET.SubElement(operator, 'Регистрационный_номер').text = result['data']['reg_number']
            ET.SubElement(operator, 'Наименование').text = name
            ET.SubElement(operator, 'ИНН').text = inn
            ET.SubElement(operator, 'Тип_оператора').text = result['data']['operator_type']
            ET.SubElement(operator, 'Основание_включения').text = result['data']['inclusion_basis']
            ET.SubElement(operator, 'Дата_регистрации').text = result['data']['registration_date']
            ET.SubElement(operator, 'Дата_начала_обработки').text = result['data']['processing_start_date']
            ET.SubElement(operator, 'Ссылка_на_карточку').text = result['data']['url']
            
            # Детальная информация
            ET.SubElement(operator, 'Ответственное_лицо').text = result['data'].get('responsible_person', 'Не указано')
            ET.SubElement(operator, 'Контактные_данные').text = result['data'].get('contact_details', 'Не указано')
        else:
            ET.SubElement(operator, 'ИНН').text = result['inn']
            ET.SubElement(operator, 'Статус').text = 'Не найден в реестре'

    xmlstr = minidom.parseString(ET.tostring(root, encoding='utf-8')).toprettyxml(indent="    ")
    with open('report.xml', 'w', encoding='utf-8') as f:
        f.write(xmlstr)

def create_excel_report(results):
    data = []
    for result in results:
        if result['data']:
            # Разделяем название и ИНН
            name_inn = result['data']['name_inn']
            name = name_inn.split('ИНН:')[0].strip()
            inn = name_inn.split('ИНН:')[1].strip() if 'ИНН:' in name_inn else ''
            
            data.append({
                'ИНН': inn,
                'Статус': 'Найден',
                'Регистрационный номер': result['data']['reg_number'],
                'Наименование': name,
                'Тип оператора': result['data']['operator_type'],
                'Основание включения': result['data']['inclusion_basis'],
                'Дата регистрации': result['data']['registration_date'],
                'Дата начала обработки': result['data']['processing_start_date'],
                'Ответственное лицо': result['data'].get('responsible_person', 'Не указано'),
                'Контактные данные': result['data'].get('contact_details', 'Не указано'),
                'Email': result['data'].get('email', ''),  # Добавляем столбец Email
                'Ссылка на карточку': result['data'].get('url', '')
            })
        else:
            data.append({
                'ИНН': result['inn'],
                'Статус': 'Не найден в реестре',
                'Регистрационный номер': '',
                'Наименование': '',
                'Тип оператора': '',
                'Основание включения': '',
                'Дата регистрации': '',
                'Дата начала обработки': '',
                'Ответственное лицо': '',
                'Контактные данные': '',
                'Email': '',
                'Ссылка на карточку': ''
            })

    # Создаем DataFrame
    df = pd.DataFrame(data)
    
    # Настройка форматирования Excel остается прежней
    with pd.ExcelWriter('report.xlsx', engine='openpyxl') as writer:
        # Записываем DataFrame в Excel
        df.to_excel(writer, index=False, sheet_name='Операторы')
        
        # Получаем рабочий лист
        worksheet = writer.sheets['Операторы']
        
        # Настраиваем формат для заголовков
        header_format = {
            'bg_color': '#4F81BD',
            'font_color': 'white',
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'border': 1
        }
        
        # Настраиваем автоширину столбцов и форматирование
        for idx, column in enumerate(worksheet.columns):
            max_length = 0
            column = [cell for cell in column]
            
            # Форматируем заголовок
            header_cell = column[0]
            header_cell.fill = PatternFill(start_color='4F81BD', end_color='4F81BD', fill_type='solid')
            header_cell.font = Font(color='FFFFFF', bold=True)
            header_cell.alignment = Alignment(wrap_text=True, vertical='top')
            header_cell.border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            # Определяем максимальную ширину столбца
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            # Устанавливаем ширину столбца
            adjusted_width = min(max_length + 2, 50)  # Ограничиваем максимальную ширину
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
            
            # Форматируем все ячейки в столбце
            for cell in column[1:]:  # Пропускаем заголовок
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
        
        # Добавляем фильтры
        worksheet.auto_filter.ref = worksheet.dimensions
        
        # Замораживаем первую строку
        worksheet.freeze_panes = 'A2'

def main():
    print("Начало работы программы")
    # Читаем список ИНН
    inn_list = read_inn_list('inn.txt')
    print(f"Загружено {len(inn_list)} ИНН из файла")
    results = []
    
    # Инициализируем драйвер
    print("Инициализация браузера...")
    driver = setup_driver()
    
    try:
        # Проверяем каждый ИНН
        for inn in inn_list:
            print(f"\nПроверка ИНН: {inn}")
            data = check_operator_status(driver, inn)
            results.append({'inn': inn, 'data': data})
            print("Ожидание перед следующим запросом...")
            time.sleep(3)
        
        # Создаем отчеты
        print("\nСоздание отчетов...")
        create_xml_report(results)
        print("XML-отчет сформирован в файле report.xml")
        create_excel_report(results)
        print("Excel-отчет сформирован в файле report.xlsx")
        
    finally:
        # Закрываем браузер
        print("Закрытие браузера...")
        driver.quit()

if __name__ == "__main__":
    main()