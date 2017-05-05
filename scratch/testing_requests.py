import requests

site_url = "http://meterdata.submetersolutions.com"
login_url = "/login.php"
file_url = "/consumption_csv.php"

username = input("Enter username: ")
password = input("Enter password: ")

# Thanks to tigerFinch @ http://stackoverflow.com/a/17633072

# Fill in your details here to be posted to the login form.
login_payload = {"txtUserName": username,
                 "txtPassword": password,
                 "btnLogin":    "Login"}
query_string = {"SiteID":   "128",
                "FromDate": "02/01/2017",
                "ToDate":   "02/28/2017",
                "SiteName": "Brimley Plaza"}

# Use 'with' to ensure the session context is closed after use.
with requests.Session() as s:
    p = s.post(site_url + login_url, data=login_payload)
    # print the html returned or something more intelligent to see if it's a successful login page.
    # print(p.text)

    # An authorised request.
    r = s.get(site_url + file_url, params=query_string)
    with open("testfile.csv", 'wb') as f:
        f.write(r.content)
    # print(r.text)
