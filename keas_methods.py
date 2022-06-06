import datetime
from utilities import xlUtiliy as xl
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys


s = Service(executable_path='../chromedriver.exe')
driver = webdriver.Chrome(service=s)
keas_url = 'http://35.183.12.229/'
keas_login_url = 'http://35.183.12.229/login/index.php'
keas_users_main_page = 'http://35.183.12.229/?redirect=0'
keas_username = 'admin'
keas_password = 'Yztek123$#'
keas_dashboard_url = 'http://35.183.12.229/my/'


# Fixture method - to open web browser
def setup():
    print('**************Setting up***********')
    # Make a full screen
    driver.maximize_window()
    # Let's wait for the browser response in general
    driver.implicitly_wait(30)
    # Navigating to the keas app website
    driver.get(keas_url)
    # Checking that we're on the correct URL address, and we're seeing correct title
    if driver.current_url == keas_url and driver.title == "Shahnas's Software Quality Assurance Testing":
        print(f'We\'re at keas homepage -- {driver.current_url}')
        print(f'We\'re seeing title message -- "Shahnas s Software Quality Assurance Testing"')
    else:
        print(f'We\'re not at the keas homepage. Check your code!')
        driver.close()
        driver.quit()


def teardown():
    print('*******************Tear down*************')
    if driver is not None:
        print(f'--------------------------------------')
        print(f'Test Completed at: {datetime.datetime.now()}')
        driver.close()
        driver.quit()
        # Make a log file with dynamic fake values
        # old_instance = sys.stdout
        # log_file = open('message.log', 'w')
        # sys.stdout = log_file
        # print(f'Email: {locators.email}\nUsername: {username}\nPassword: {locators.new_password}\n'
        #       f'Full Name: {locators.full_name}')
        # sys.stdout = old_instance
        # log_file.close()


def log_in(username, password):
    print('*******************Log in*************')
    if driver.current_url == keas_url:
        driver.find_element(By.LINK_TEXT, 'Log in').click()
        if driver.current_url == keas_login_url:
            driver.find_element(By.ID, 'username').send_keys(username)
            sleep(0.25)
            driver.find_element(By.ID, 'password').send_keys(password)
            sleep(0.25)
            driver.find_element(By.ID, 'loginbtn').click()
            if driver.title == 'Dashboard' and driver.current_url == keas_dashboard_url:
                assert driver.current_url == keas_dashboard_url
                print(f'Log in successfully. Dashboard is present. \n'
                      f'We logged in with Username: {username} at {datetime.datetime.now()}')
            else:
                print(f'We\re not at the Dashboard. Try again')


def log_out():
    print('*******************Log Out*************')
    driver.find_element(By.CLASS_NAME, 'userpicture').click()
    sleep(0.25)
    driver.find_element(By.XPATH, '//span[contains(., "Log out")]').click()
    sleep(0.25)
    if driver.current_url == keas_url:
        print(f'Log out successfully at: {datetime.datetime.now()}')


def add_new_user(file, sheetname):
    print('*******************Add new user*************')
    # path = "C:/Users/kando/Desktop/YZTek/data_keas.xlsx"
    rows = xl.get_row_count(file, sheetname)
    for r in range(2, rows+1):
        driver.find_element(By.XPATH, '//span[contains(., "Site administration")]').click()
        sleep(0.25)
        assert driver.find_element(By.LINK_TEXT, 'Users').is_displayed()
        driver.find_element(By.LINK_TEXT, 'Users').click()
        sleep(0.25)
        driver.find_element(By.LINK_TEXT, 'Add a new user').click()
        sleep(0.25)
        assert driver.find_element(By.LINK_TEXT, 'Add a new user').is_displayed()
        sleep(0.25)
    # Enter fake data into username open field
        username = xl.read_data(file, sheetname, r, 1)
        driver.find_element(By.ID, 'id_username').send_keys(username)
        sleep(0.25)
        # Click by the password open field and enter fake password
        driver.find_element(By.LINK_TEXT, 'Click to enter text').click()
        sleep(0.25)
        password = xl.read_data(file, sheetname, r, 2)
        driver.find_element(By.ID, 'id_newpassword').send_keys(password)
        sleep(0.25)
        first_name = xl.read_data(file, sheetname, r, 3)
        driver.find_element(By.ID, 'id_firstname').send_keys(first_name)
        sleep(0.25)
        last_name = xl.read_data(file, sheetname, r, 4)
        driver.find_element(By.ID, 'id_lastname').send_keys(last_name)
        sleep(0.25)
        email = xl.read_data(file, sheetname, r, 5)
        driver.find_element(By.ID, 'id_email').send_keys(email)
        # Select 'Allow everyone to see my email address'
        Select(driver.find_element(By.ID, 'id_maildisplay')).select_by_visible_text(
            'Allow everyone to see my email address')
        sleep(0.25)
        keas_net_profile = f'https://keas.net/{username}'
        driver.find_element(By.ID, 'id_moodlenetprofile').send_keys(keas_net_profile)
        sleep(0.25)
        city = xl.read_data(file, sheetname, r, 6)
        driver.find_element(By.ID, 'id_city').send_keys(city)
        sleep(0.25)
        Select(driver.find_element(By.ID, 'id_country')).select_by_visible_text('Canada')
        sleep(0.25)
        Select(driver.find_element(By.ID, 'id_timezone')).select_by_visible_text('America/Vancouver')
        sleep(0.25)
        driver.find_element(By.ID, 'id_description_editoreditable').clear()
        sleep(0.25)
        # driver.find_element(By.ID, 'id_description_editoreditable').send_keys(description)
        description = f'User added by Automation engineer {keas_username}via Python Selenium ' \
                      f'Automated script on {datetime.datetime.now()}'
        driver.find_element(By.ID, 'id_description_editoreditable').send_keys(description)
        # Upload picture to the User Picture section
        # Click by 'You can drag and drop files here to add them.' section
        driver.find_element(By.XPATH, "//div[@class ='fp-btn-add']//i[@class ='icon fa fa-file-o fa-fw ']").click()

        sleep(1)
        driver.find_element(By.PARTIAL_LINK_TEXT, 'Server files').click()
        sleep(1)
        driver.find_element(By.PARTIAL_LINK_TEXT, 'IT').click()
        sleep(1)
        driver.find_element(By.PARTIAL_LINK_TEXT, 'Quality Assurance').click()
        sleep(1)
        driver.find_element(By.PARTIAL_LINK_TEXT, 'Course image').click()
        sleep(1)
        driver.find_element(By.PARTIAL_LINK_TEXT, 'wallpaper1.jpg').click()
        sleep(1)
        # Click by 'Select this file' button
        driver.find_element(By.XPATH, '//button[contains(., "Select this file")]').click()
        sleep(1)
        # Enter value to the 'Picture description' open field
        driver.find_element(By.ID, 'id_imagealt').send_keys(first_name)
        sleep(0.25)
        # Click by 'Additional names' dropdown menu
        driver.find_element(By.XPATH, '//a[contains(., "Additional names")]').click()
        driver.find_element(By.ID, 'id_firstnamephonetic').send_keys(first_name)
        driver.find_element(By.ID, 'id_lastnamephonetic').send_keys(last_name)
        driver.find_element(By.ID, 'id_middlename').send_keys(" ")
        driver.find_element(By.ID, 'id_alternatename').send_keys(first_name)
        sleep(0.25)
        # Click by 'Interests' dropdown menu
        driver.find_element(By.XPATH, '//a[contains(., "Interests")]').click()
        sleep(0.25)

        # Using for loop, take all items from the list and populate data
        list_of_interests = [username, password, first_name, email, city]
        for tag in list_of_interests:
            driver.find_element(By.XPATH, '//div[3]/input').click()
            sleep(0.25)
            driver.find_element(By.XPATH, '//div[3]/input').send_keys(tag)
            sleep(0.25)
            driver.find_element(By.XPATH, '//div[3]/input').send_keys(Keys.ENTER)
        sleep(0.25)
        # Fill Optional section
        # Click by Optional link to open that section
        driver.find_element(By.XPATH, "//a[text() = 'Optional']").click()
        sleep(0.25)
        # Fill out the Web page input open field
        driver.find_element(By.CSS_SELECTOR, "input#id_url").send_keys(f'www.{username}.com')
        sleep(0.25)
        # Fill out the ICQ Number input open field
        driver.find_element(By.CSS_SELECTOR, "input#id_icq").send_keys(123456789)
        sleep(0.25)
        # Fill out the Skype ID input open field
        driver.find_element(By.CSS_SELECTOR, "input#id_skype").send_keys(username)
        sleep(0.25)
        # Fill out the AIM ID input open field
        driver.find_element(By.CSS_SELECTOR, "input#id_aim").send_keys(username)
        sleep(0.25)
        # Fill out the Yahoo ID input open field
        driver.find_element(By.CSS_SELECTOR, "input#id_yahoo").send_keys(username)
        sleep(0.25)
        # Fill out the MSN ID input open field
        driver.find_element(By.CSS_SELECTOR, "input#id_msn").send_keys(username)
        sleep(0.25)
        # Fill out the ID number input open field
        driver.find_element(By.CSS_SELECTOR, "input#id_idnumber").send_keys(123456789)
        sleep(0.25)
        # Fill out the Institution input open field
        driver.find_element(By.CSS_SELECTOR, "input#id_institution").send_keys('CCTB')
        sleep(0.25)
        # Fill out the Department input open field
        driver.find_element(By.CSS_SELECTOR, "input#id_department").send_keys('SQTA')
        sleep(0.25)
        # Fill out the Phone input open field
        phone = xl.read_data(file, sheetname, r, 7)
        driver.find_element(By.CSS_SELECTOR, "input#id_phone1").send_keys(phone)
        sleep(0.25)
        # Fill out the Mobile Phone input open field
        driver.find_element(By.CSS_SELECTOR, "input#id_phone2").send_keys(phone)
        sleep(0.25)
        # Fill out the Address input open field
        address = xl.read_data(file, sheetname, r, 8)
        driver.find_element(By.CSS_SELECTOR, "input#id_address").send_keys(address)
        sleep(0.25)
        # Click by 'Create user' button
        driver.find_element(By.ID, 'id_submitbutton').click()
        sleep(0.25)
        print(f'---New user created with username: "{username}" at:{datetime.datetime.now()}')


# setup()
# log_in(keas_username, keas_password)
# create_new_user()
# add_new_user("C:/Users/kando/Desktop/YZTek/data_keas.xlsx", "Sheet1")
# log_out()
# tearDown()
