from sys import stdout
from requests import Session
from datetime import datetime, timedelta
from getpass import getpass

site_url = "http://meterdata.submetersolutions.com"
login_url = "/login.php"
file_url = "/consumption_csv.php"
terminal = stdout.isatty()  # not all functions work on PyCharm


def get_data(site_id, site_name, period=None):
    """
    Access the online submeter database to download and save
    data for a given (or asked) period.
    Requires authentication.

    :param str site_id: the looked-up "SiteID" param in the data query string
    :param str site_name: the "SiteName" param in the data query string
    :param str period: the month(s) to get data for
    :return:
    """
    username = input("Username: ")
    password = getpass() if terminal else input("Enter password: ")

    # Get period to process (if not given)
    period = period or input("Enter a period to get data for: ")
    periods = []
    months = 0
    try:
        if len(period) == 7:  # one month
            start = datetime.strptime(period, "%b%Y")
            end = last_day_of_month(start)
            periods.append((start, end))
            months += 1
        else:  # a period
            first = datetime.strptime(period[:7], "%b%Y")
            last = datetime.strptime(period[-7:], "%b%Y")
            months += (last.year - first.year)*12 + last.month - first.month + 1
            start = first
            for _ in range(months):
                end = last_day_of_month(start)
                periods.append((start, end))
                start = next_month(start)
    except ValueError as e:
        raise Exception("Incorrect period format. Accepted formats:\n"
                        "\tJan2016         (single month)\n"
                        "\tJan2016-Feb2017 (range of months)") from e

    # print(*periods, sep="\n")

    # (Thanks to tigerFinch @ http://stackoverflow.com/a/17633072)
    # Fill in your details here to be posted to the login form.
    login_payload = {"txtUserName": username,
                     "txtPassword": password,
                     "btnLogin": "Login"}
    query_string = {"SiteID": site_id,
                    "SiteName": site_name}
    # print(query_string)

    # Use 'with' to ensure the session context is closed after use.
    with Session() as session:
        # response = session.post(site_url + login_url, data=login_payload)
        # assert response.status_code == 200
        # this is true even if user/pass is incorrect, so:
        # TODO: find a way to verify correct user/pass

        update_progress_bar(0)  # start progress bar
        for idx, (start, end) in enumerate(periods):
            # period = start.strftime("Data/%b%Y_data.csv")
            # query_string["FromDate"] = start.strftime("%m/%d/%y")
            # query_string["ToDate"] = end.strftime("%m/%d/%y")
            # # print(query_string, '\n')
            #
            # # An authorised request.
            # response = session.get(site_url + file_url, params=query_string)
            # assert response.status_code == 200
            # with open(period, 'xb') as f:
            #     f.write(response.content)
            update_progress_bar((idx+1) / months)
    print("Data download complete. See 'Data' folder for files.")


def next_month(date):
    month_after = date.replace(day=28) + timedelta(days=4)  # never fails
    return month_after.replace(day=1)


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
    month_after = next_month(date)
    return month_after - timedelta(days=month_after.day)


def update_progress_bar(percent: float):
    if not terminal:  # because PyCharm doesn't treat '\r' well
        print("[{}{}]".format('#' * int(percent * 20),
                              ' ' * (20 - int(percent * 20))))
    elif percent == 1:
        print("Progress: {:3.1%}".format(percent))
    else:
        print("Progress: {:3.1%}\r".format(percent), end="")


if __name__ == "__main__":
    if not terminal:
        print("WARNING: This is not a TTY/terminal. "
              "Passwords will not be hidden.")
    get_data("128", "Brimley Plaza")