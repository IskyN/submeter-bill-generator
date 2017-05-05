from selenium.webdriver import Chrome  # , PhantomJS
from time import sleep
from datetime import datetime, timedelta

url = "http://meterdata.submetersolutions.com/"


def get_data(driver):
    driver.set_window_size(1440, 960)
    driver.get(url + "login.php")

    # Find username and password entry boxes, and login button
    un_box = driver.find_element_by_id("txtUserName")
    pw_box = driver.find_element_by_id("txtPassword")
    login_button = driver.find_element_by_id("btnLogin")

    username = input("Enter username: ")
    password = input("Enter password: ")
    un_box.clear()
    un_box.send_keys(username)
    pw_box.clear()
    pw_box.send_keys(password)
    login_button.click()
    sleep(1)
    while driver.find_elements_by_class_name("Error"):  # ensure proper login
        print("Incorrect username/password. Please try again.")
        username = input("Enter username: ")
        password = input("Enter password: ")
        un_box.clear()
        un_box.send_keys(username)
        pw_box.clear()
        pw_box.send_keys(password)
        login_button.click()
        sleep(1)

    # Get month to process
    month = input("Enter a month/year to get data for. "
                  "Format: Jan. 2016: ")
    start_date = datetime.strptime(month, "%b. %Y")
    end_date = last_day_of_month(start_date)
    driver.find_element_by_id("objPropertyList_1_0").send_keys(username)
    sleep(1)
    start_box = driver.find_element_by_id("R1_fromdate")
    end_box = driver.find_element_by_id("R1_todate")
    start_box.clear()
    start_box.send_keys(start_date.strftime("%m/%d/%Y"))
    end_box.clear()
    end_box.send_keys(end_date.strftime("%m/%d/%Y"))
    driver.find_element_by_id("R1_Refresh").click()
    sleep(1)
    driver.find_element_by_id("R1_Export").click()
    sleep(1)

    # Close Consumption dialog box
    driver.find_element_by_class_name("ui-dialog-titlebar-close").click()
    # Log out
    driver.get(url + "logout.php")


def last_day_of_month(date):
    """
    Return the last day of the given month (leap year-sensitive),
    with date unchanged.
    Thanks to Augusto Men: http://stackoverflow.com/a/13565185

    :param datetime date: the first day of the given month
    :return: datetime

    >>> d = datetime(2012, 2, 1)
    >>> last_day_of_month(d)
    datetime.datetime(2012, 2, 29, 0, 0)
    >>> d.day == 1
    True
    """
    next_month = date.replace(day=28) + timedelta(days=4)  # never fails
    return next_month - timedelta(days=next_month.day)


if __name__ == "__main__":
    # driver = PhantomJS()
    driver = Chrome()
    try:
        get_data(driver)
    finally:
        driver.quit()