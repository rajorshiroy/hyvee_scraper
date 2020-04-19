import csv
import os
import re
import time
import json
import glob
import shutil
import zipfile
import getpass
import requests
import subprocess
import multiprocessing
from pprint import pprint
from config import config
from threading import Thread
from zipfile import BadZipFile
from selenium.webdriver.common.by import By
from urllib3.exceptions import ProtocolError
from appium import webdriver as appium_driver
from selenium import webdriver as selenium_driver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import NoSuchElementException, WebDriverException


class HyVee:
    def __init__(self):
        self.driver = None
        self.driver_wait = None
        self.current_page = None
        self.item_names = None
        self.session = None
        self.existing_items = None
        self.search_form_data = None
        self.stores = []
        self.username = config.email
        self.password = config.password

    def init_driver(self):
        chrome_options = selenium_driver.ChromeOptions()
        self.driver = selenium_driver.Chrome(options=chrome_options)

    def login(self):
        if not self.driver:
            self.init_driver()
        self.driver.get('https://www.hy-vee.com/my-hy-vee/account_preferences.aspx')
        try:
            # enter username and password
            WebDriverWait(self.driver, 10).until(
                ec.presence_of_element_located((By.ID, 'username'))).send_keys(self.username)
            WebDriverWait(self.driver, 10).until(
                ec.presence_of_element_located((By.ID, 'password'))).send_keys(self.password)

            # log current url
            current_url = self.driver.current_url

            # click on log in button
            WebDriverWait(self.driver, 10).until(
                ec.presence_of_element_located((By.XPATH, '//button[@label="Log In"]'))).click()
        except:
            print('failed to login. retrying...')
            self.login()

        # check for url change
        while self.driver.current_url == current_url:
            time.sleep(0.5)

        # if changed url is not my-account then login failed
        if not self.driver.current_url == 'https://www.hy-vee.com/my-account':
            print('failed to login. retrying...')
            self.login()

    def search_stores(self):
        print('searching stores...')
        print('-' * 75)

        if not self.driver:
            self.init_driver()
        self.driver.get(
            'https://www.hy-vee.com/stores/store-finder-results.aspx?zip=&state=&city=&olfloral=False&olcatering=False&olgrocery=False&olpre=False&olbakery=False&diet=False&chef=False')

        while True:
            self.current_page = self.driver.find_element_by_class_name('current_page').text

            table = self.driver.find_element_by_id('ctl00_cph_main_content_spuStoreFinderResults_gvStores')
            table_rows = table.find_elements_by_tag_name('tr')[:-2]
            for row in table_rows:
                td = row.find_elements_by_tag_name('td')[-1]
                a_tag = td.find_element_by_tag_name('a')
                store_id = a_tag.get_attribute('storeid')
                store_name = a_tag.text.replace('#', '')
                tag_id = a_tag.get_attribute('id')
                script = f'return document.querySelector("#{tag_id}").parentElement.innerText;'
                inner_text = self.driver.execute_script(script).split('\n')

                if 'Hy-Vee' in inner_text[1]:
                    address = inner_text[2]
                else:
                    address = inner_text[1]

                self.stores.append({
                    'id': store_id,
                    'name': store_name,
                    'address': address
                })

                print('name:', store_name)
                print('store_id:', store_id)
                print('address:', address)
                print('-' * 75)

            try:
                next_btn = self.driver.find_element_by_id(
                    'ctl00_cph_main_content_spuStoreFinderResults_gvStores_ctl10_btnNext')
            except NoSuchElementException:
                break

            if next_btn.get_attribute('class'):
                break
            else:
                next_btn.click()
                self.wait_till_next_page_loads()
        print(f'{len(self.stores)} stores found')
        self.save_stores(self.stores)

    def change_store(self, store_id):
        self.driver.get(f'https://www.hy-vee.com/my-hy-vee/account_preferences.aspx?s={store_id}')
        while not self.driver.current_url == 'https://www.hy-vee.com/my-account':
            time.sleep(0.5)
        time.sleep(1)

    def wait_till_next_page_loads(self):
        while True:
            current_page = self.driver.find_element_by_class_name('current_page').text
            if current_page != self.current_page:
                return
            else:
                time.sleep(0.5)

    @staticmethod
    def save_stores(data):
        with open('site_cache/stores.json', 'w', encoding='utf-8') as f:
            f.write(json.dumps(data, indent=2))

    def close_driver(self):
        self.driver.close()

    def get_items(self):
        if os.path.exists('inputs/items.csv'):
            with open('inputs/items.csv', encoding='utf-8') as f:
                self.item_names = f.read().strip().split('\n')
        else:
            print('items.csv not found in inputs folder...')
            quit()

    def get_stores_from_file(self):
        if not os.path.exists('site_cache/stores.json'):
            print('stores.json not found in site_cache. searching for stores...')
            hyvee.search_stores()
            hyvee.close_driver()
        else:
            with open('site_cache/stores.json') as f:
                self.stores = json.loads(f.read())

    def visit_list_page(self):
        self.driver.get('https://www.hy-vee.com/deals/shopping-list.aspx')
        time.sleep(3)
        self.session = requests.session()

        # get the cookies from this page
        self.session.cookies.update({cookie['name']: cookie['value'] for cookie in self.driver.get_cookies()})

        # add the user agent to session header
        self.session.headers = {
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Safari/537.36'
        }

        # create the form data using data found in this page
        self.search_form_data = {
            '__EVENTTARGET': 'ctl00$ContentPlaceHolder1$btnAddItem',
            '__VIEWSTATE': self.driver.find_element_by_id('__VIEWSTATE').get_attribute('value'),
            '__VIEWSTATEGENERATOR': self.driver.find_element_by_id('__VIEWSTATEGENERATOR').get_attribute('value'),
            '__VIEWSTATEENCRYPTED': self.driver.find_element_by_id('__VIEWSTATEENCRYPTED').get_attribute('value'),
            '__EVENTVALIDATION': self.driver.find_element_by_id('__EVENTVALIDATION').get_attribute('value'),
            'ctl00$ContentPlaceHolder1$txtAddItem': '',
        }

        view_state_index = 0
        while True:
            try:
                view_state_index += 1
                self.search_form_data[f'__VIEWSTATE{view_state_index}'] = self.driver.find_element_by_id(
                    f'__VIEWSTATE{view_state_index}').get_attribute('value')
            except:
                break
        self.search_form_data['__VIEWSTATEFIELDCOUNT'] = view_state_index

        # get existing list items
        item_cells = self.driver.find_elements_by_class_name('cellDescription')
        self.existing_items = [cell.text.strip() for cell in item_cells]

    def add_item_to_list(self, item_name, index=''):
        if item_name in self.existing_items:
            print(f'{index}{item_name}: already added')
            return

        data = self.search_form_data.copy()
        data['ctl00$ContentPlaceHolder1$txtAddItem'] = item_name

        while True:
            try:
                response = self.session.post('https://www.hy-vee.com/deals/shopping-list.aspx', data=data)
                break
            except:
                print('connection broken. retrying...')
                time.sleep(1)

        if response.status_code == 200:
            print(f'{index}{item_name}: added')
        else:
            print(f'{index}{item_name}: failed to add. status code: {response.status_code}')
            print('session headers:')
            pprint(self.session.headers)

            print('session cookies')
            pprint(dict(self.session.cookies))

            print('form data:')
            pprint(data)
            quit()

    def remove_list_items(self, count):
        remove_tags = self.driver.find_elements_by_class_name('listRemove')
        remove_scripts = [tag.get_attribute('href')[11:].strip() for tag in remove_tags]

        if count != 'all':
            try:
                count = int(count)
            except:
                print('enter a valid count. example: all, 5, 10, 23, 100')
                quit()
            remove_scripts = remove_scripts[:count]

        for i, script in enumerate(remove_scripts):
            self.driver.execute_script(script)
            print(f'removed {i + 1} item{"" if i == 0 else "s"}')


class Appium:
    def __init__(self):
        self.driver = None
        self.nox_process = None

    def init_driver(self):
        desired_cap = {
            'platformName': 'Android',
            'platformVersion': config.android_version,
            'deviceName': config.android_device_name,
            'noReset': True,
            'appPackage': 'com.hyvee.android',
            'appActivity': 'com.hyvee.android.ui.MainActivity'
        }

        self.driver = appium_driver.Remote(f'http://localhost:{config.appium_port}/wd/hub',
                                           desired_capabilities=desired_cap)

        time.sleep(5)

    def connect_to_nox(self):
        self.nox_process = multiprocessing.Process(target=subprocess.run,
                                                   args=(f"adb connect 127.0.0.1:{config.nox_port}", True))
        self.nox_process.start()

    def open_list(self):
        while True:
            try:
                # print('opening list page in Hy-Vee app')
                self.driver.find_element_by_id('com.hyvee.android:id/bottom_nav_list').click()
                time.sleep(2)
                return
            except (WebDriverException, ProtocolError):
                self.init_driver()

    def open_deals(self):
        while True:
            try:
                # print('opening deals page in Hy-Vee app')
                self.driver.find_element_by_id('com.hyvee.android:id/bottom_nav_deals').click()
                time.sleep(2)
                return
            except (WebDriverException, ProtocolError):
                self.init_driver()

    def open_my_account(self):
        while True:
            try:
                # print('opening deals page in Hy-Vee app')
                self.driver.find_element_by_id('com.hyvee.android:id/bottom_nav_more').click()
                time.sleep(1)
                self.driver.find_element_by_android_uiautomator(
                    'new UiSelector().textContains("My Account")').click()
                time.sleep(2)
                return
            except (WebDriverException, ProtocolError):
                self.init_driver()

    @staticmethod
    def open_appium():
        subprocess.run(config.appium_location, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    def open_appium_threaded(self):
        appium_thread = Thread(target=self.open_appium)
        appium_thread.start()

    @staticmethod
    def close_appium():
        subprocess.run('taskkill /IM Appium.exe /F', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


class Fiddler:
    def __init__(self):
        self.username = getpass.getuser()
        self.fiddler_archive_file = f'C:\\Users\\{self.username}\\Documents\\Fiddler2\\Captures\\dump.saz'
        self.sessions = []
        self.fiddler_process = None

    def open_fiddler(self):
        self.fiddler_process = multiprocessing.Process(target=subprocess.run, args=('fiddler', True))
        self.fiddler_process.start()

    @staticmethod
    def save_fiddler_session():
        subprocess.run('execaction dump', shell=True)

    @staticmethod
    def close_fiddler():
        subprocess.run('taskkill /IM Fiddler.exe', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    def unpack_saz(self):
        while not os.path.exists(self.fiddler_archive_file):
            print('waiting for fiddler dump...')
            time.sleep(1)
        while True:
            try:
                with zipfile.ZipFile(self.fiddler_archive_file, 'r') as zip_ref:
                    zip_ref.extractall('fiddler_session')
                return
            except BadZipFile:
                print('waiting for saz compilation...')
                time.sleep(1)

    @staticmethod
    def clean_fiddler_session():
        shutil.rmtree('fiddler_session', ignore_errors=True)

    def clean_dump(self):
        if os.path.exists(self.fiddler_archive_file):
            os.remove(self.fiddler_archive_file)

    def parse_sessions(self):
        pass


class SazParser:
    def __init__(self):
        self.stores_found = []

    def get_requests(self):
        os.makedirs('site_cache/aisles/', exist_ok=True)

        request_files = glob.glob('fiddler_session\\raw\\*_c.txt')
        for request_file in request_files:
            with open(request_file) as f:
                data = f.read()
            if 'https://api.hy-vee.com/ShoppingLists/' in data:
                # get the store id from requests file
                store_id = re.findall(r'(?<=/items/).*?(?= HTTP)', data)
                if len(store_id) == 0:
                    continue
                else:
                    store_id = store_id[0]
                    if store_id in self.stores_found:
                        continue
                # get the response filename
                response_file = request_file.replace('_c.txt', '_s.txt')

                # get the response json
                with open(response_file) as f:
                    response_data = f.read()

                try:
                    response_json = response_data[response_data.index('\n{'):].strip()
                except ValueError:
                    continue
                response_json = json.loads(response_json)

                # parse json to get aisle numbers
                aisle_numbers = {}
                try:
                    shopping_list = response_json['data']['shopping_list_item_list']
                except KeyError:
                    continue

                for item in shopping_list:
                    if item['aisle']:
                        aisle_numbers[item['item_base']['description']] = item['aisle']

                if len(aisle_numbers) > 0:
                    self.stores_found.append(store_id)

                if os.path.exists(f'site_cache/aisles/{store_id}.json'):
                    with open(f'site_cache/aisles/{store_id}.json') as f:
                        aisle_data = json.loads(f.read())
                    aisle_data.update(aisle_numbers)
                else:
                    aisle_data = aisle_numbers

                with open(f'site_cache/aisles/{store_id}.json', 'w', encoding='utf-8') as f:
                    f.write(json.dumps(aisle_data, indent=2))


class FileHandler:
    def __init__(self):
        self.completed_stores = []
        self.skip_completed = True
        self._store_filenames = {}
        self._get_completed_stores()
        self._get_store_filenames()

    def _get_store_filenames(self):
        with open('site_cache/stores.json') as f:
            stores = json.loads(f.read())
        for store in stores:
            self._store_filenames[str(store["id"])] = f'{store["name"]}_{store["id"]}.csv'

    def get_store_filename(self, store_id):
        return self._store_filenames.get(store_id)

    def _get_completed_stores(self):
        stores = glob.glob('aisle_data/*.csv')
        for store in stores:
            store_id = store.split('.')[-2].split('_')[-1]
            self.completed_stores.append(store_id)
        if len(self.completed_stores) > 0:
            self.skip_completed = (f'{self.completed_stores} stores found under aisle_data.'
                                   f' do you want to skip them? [ y/n ]').lower() == 'y'

    def save_scraped_data(self, store_id_list):
        for store_id in store_id_list:
            try:
                with open(f'site_cache/aisles/{store_id}.json') as f:
                    data = json.loads(f.read())
            except FileNotFoundError:
                continue
            csv_data = [[item, isle] for item, isle in data.items()]

            filename = self.get_store_filename(store_id)
            with open(f'aisle_data/{filename}', 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                for row in csv_data:
                    writer.writerow(row)
            print(f'saved {filename} in aisle_data')


class Importer:
    def __init__(self):
        self.session = requests.session()
        self.import_session = requests.session()
        self.files = None
        self.form_body_template = None

    def login(self):
        print('logging in to speedshopperapp dashboard...')

        self.session.headers = {
            'Referer': 'https://www.speedshopperapp.com/app/admin/login',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36'
        }

        payload = {
            'username': config.import_username,
            'password': config.import_password
        }

        response = self.session.post('https://www.speedshopperapp.com/app/admin/login', data=payload)

        if '<title>Dashboard</title>' in response.text:
            print('logged in')
            self.import_session.headers = {
                'Host': 'www.speedshopperapp.com',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:70.0) Gecko/20100101 Firefox/70.0',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'en-US,en;q=0.5',
                'Accept-Encoding': 'gzip, deflate, br',
                'Content-Type': 'multipart/form-data; boundary=---------------------------1267546269709',
                'Origin': 'https://www.speedshopperapp.com',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1',
            }
            self.import_session.cookies = self.session.cookies
        else:
            print('login failed. check your username and password in import_user.json file')
            quit()

    def get_files(self):
        self.files = glob.glob('aisle_data/*.csv')
        if len(self.files) == 0:
            print('no csv files found under aisle_data folder')
            quit()
        count = len(self.files)
        print(f'{count} {"file" if count == 1 else "files"} found')

    def search_store(self, address='', name=''):
        """searches the website for all available stores and saves their details for quick access"""

        data = {
            'draw': '2',
            'columns[0][data]': '0',
            'columns[0][name]': '',
            'columns[0][searchable]': 'true',
            'columns[0][orderable]': 'true',
            'columns[0][search][value]': '',
            'columns[0][search][regex]': 'false',
            'columns[1][data]': '1',
            'columns[1][name]': '',
            'columns[1][searchable]': 'true',
            'columns[1][orderable]': 'true',
            'columns[1][search][value]': '',
            'columns[1][search][regex]': 'false',
            'columns[2][data]': '2',
            'columns[2][name]': '',
            'columns[2][searchable]': 'true',
            'columns[2][orderable]': 'true',
            'columns[2][search][value]': '',
            'columns[2][search][regex]': 'false',
            'columns[3][data]': '3',
            'columns[3][name]': '',
            'columns[3][searchable]': 'true',
            'columns[3][orderable]': 'true',
            'columns[3][search][value]': '',
            'columns[3][search][regex]': 'false',
            'columns[4][data]': '4',
            'columns[4][name]': '',
            'columns[4][searchable]': 'true',
            'columns[4][orderable]': 'true',
            'columns[4][search][value]': '',
            'columns[4][search][regex]': 'false',
            'columns[5][data]': '5',
            'columns[5][name]': '',
            'columns[5][searchable]': 'true',
            'columns[5][orderable]': 'true',
            'columns[5][search][value]': '',
            'columns[5][search][regex]': 'false',
            'columns[6][data]': '6',
            'columns[6][name]': '',
            'columns[6][searchable]': 'true',
            'columns[6][orderable]': 'true',
            'columns[6][search][value]': '',
            'columns[6][search][regex]': 'false',
            'columns[7][data]': '7',
            'columns[7][name]': '',
            'columns[7][searchable]': 'true',
            'columns[7][orderable]': 'true',
            'columns[7][search][value]': '',
            'columns[7][search][regex]': 'false',
            'columns[8][data]': '8',
            'columns[8][name]': '',
            'columns[8][searchable]': 'true',
            'columns[8][orderable]': 'false',
            'columns[8][search][value]': '',
            'columns[8][search][regex]': 'false',
            'order[0][column]': '0',
            'order[0][dir]': 'asc',
            'start': '0',
            'length': '50000',
            'search[value]': '',
            'search[regex]': 'false',
            'name': name,
            'address': address
        }

        res = self.session.post('https://www.speedshopperapp.com/app/admin/stores/getstores', data=data)

        try:
            data = res.json()['data'][0]
            return re.search(r'[0-9]+', data[4]).group()
        except IndexError:
            return None

    def import_file(self, file_path, filename, store_id):
        """imports a csv file in the website"""
        response = self.import_session.post('https://www.speedshopperapp.com/app/admin/stores/importFile',
                                            data=self.get_form_body(file_path, filename, store_id))

        if 'Imported items successfully' in response.text:
            return True
        return False

    def get_form_body(self, file_path, file_name, store_id):
        if not self.form_body_template:
            if os.path.exists('config/request-body.txt'):
                with open('config/request-body.txt') as f:
                    self.form_body_template = f.read().strip()
            else:
                print('request-body.txt not found')
                quit()

        with open(file_path, encoding='utf-8') as f:
            data = f.read().strip()

        # remove items with "char" in name
        data = '\n'.join([line for line in data.split('\n') if 'char' not in line.lower()])

        body = self.form_body_template % (config.import_id, file_name, data, store_id)
        print('-' * 100)
        print(body)
        print('-' * 100)
        return body.encode('utf-8')


class Address:
    def __init__(self):
        self.store_addresses = {}
        self.get_store_addresses()

    def get_address(self, store_id):
        return self.store_addresses.get(store_id)

    def get_store_addresses(self):
        if os.path.exists('site_cache/stores.json'):
            with open('site_cache/stores.json') as f:
                stores = json.loads(f.read())
                for store in stores:
                    self.store_addresses[store['id']] = store['address']
        else:
            print('stores.json not found in site_cache folder')
            quit()


if __name__ == '__main__':
    hyvee = HyVee()
    print('choose an option:')
    print('1. add items to list')
    print('2. remove items from list')
    print('3. search stores')
    print('4. get aisles')
    print('5. import csv')
    option = input('option: ').strip()

    if option == '1':
        hyvee.get_items()
        print('logging in...')
        hyvee.login()
        print('logged in.')

        print('getting existing list...')
        hyvee.visit_list_page()
        hyvee.close_driver()
        print('driver closed.')

        print('adding items to list...')
        print('-' * 75)

        # calculate the items to be added to list using input from config file
        li, ui = min(config.items_to_add), max(config.items_to_add)
        lower_index = max(li - 1, 0)
        upper_index = min(ui - 1, len(hyvee.item_names) - 1, lower_index + 99)
        items_to_add = hyvee.item_names[lower_index:upper_index]

        for i, item in enumerate(items_to_add):
            if item.strip() == '':
                continue
            hyvee.add_item_to_list(item, index=f'[{i + 1} of {len(items_to_add)}] ')

    elif option == '2':
        rem_count = input('how many items to remove? [ all, 5, 10, 23 etc. ]: ')
        print('logging in...')
        hyvee.login()
        print('logged in.')

        print('getting existing list...')
        hyvee.visit_list_page()
        print(f'deleting {rem_count} items from list')
        hyvee.remove_list_items(rem_count)
        hyvee.close_driver()
        print('driver closed.')

    elif option == '3':
        hyvee.search_stores()
        hyvee.close_driver()

    elif option == '4':
        appium = Appium()
        fiddler = Fiddler()
        file_handler = FileHandler()

        # start appium and fiddler
        print('starting appium...')
        appium.open_appium_threaded()
        print('starting fiddler...')
        fiddler.open_fiddler()

        print('-' * 75)
        print('follow these steps and hit enter once done...')
        print('* start and unlock the android emulator')
        print('* start the appium server')
        print('-' * 75)
        input('')

        print('connecting to Nox...')
        appium.connect_to_nox()

        print('logging in to Hy-vee...')
        hyvee.login()
        print('logged in!')
        print('visiting list page...')
        hyvee.visit_list_page()
        hyvee.get_stores_from_file()

        print('opening HyVee app...')
        appium.init_driver()

        # choose how many stores to scrape
        if config.stores_to_scrape == 'all':
            stores_to_scrape = hyvee.stores
        else:
            try:
                stores_to_scrape = hyvee.stores[:int(config.stores_to_scrape)]
            except:
                stores_to_scrape = hyvee.stores

        if file_handler.skip_completed:
            stores_to_scrape = [store for store in stores_to_scrape if store not in file_handler.completed_stores]

        for i, store in enumerate(stores_to_scrape):
            print(f'[{i + 1} of {len(stores_to_scrape)}] changing store to: {store["name"]} [ {store["id"]} ]')
            hyvee.change_store(store['id'])
            appium.open_my_account()
            appium.open_list()

        hyvee.close_driver()

        fiddler.save_fiddler_session()

        # terminate appium thread
        appium.nox_process.terminate()

        fiddler.unpack_saz()
        fiddler.close_fiddler()
        appium.close_appium()

        # terminate fiddler Thread
        fiddler.fiddler_process.terminate()

        # parse the saz file and remove fiddler session
        saz_parser = SazParser()
        saz_parser.get_requests()
        fiddler.clean_fiddler_session()

        # save the csv files
        file_handler.save_scraped_data(saz_parser.stores_found)

    elif option == '5':
        address = Address()
        importer = Importer()
        importer.login()
        importer.get_files()
        print('-' * 75)
        for file in importer.files:
            filename = file.split('\\')[-1]
            print(f'filename: {filename}')
            store_id = filename[:-4].split('_')[-1]
            print(f'store id: {store_id}')
            street_address = address.get_address(store_id=store_id)
            print(f'street address: {street_address}')
            site_id = importer.search_store(address=street_address)
            if not site_id:
                print('store not found on speedshopperapp')
                print('-' * 75)
                continue
            print(f'import url: https://www.speedshopperapp.com/app/admin/stores/import/{site_id}')
            success = importer.import_file(file, filename, site_id)
            if success:
                print('imported successfully')
            else:
                print('failed to import')
            print('-' * 75)

    else:
        print('choose a valid option')
