from pandas import read_excel
from datetime import datetime
from smtplib import SMTP
import json


def send_email(name, email, msg):
    with open('config.json', 'r') as c:
        params = json.load(c)["params"]
            
    GMAIL_ID = params["Gmail_username"]
    GMAIL_PASWD = params["Gmail_password"]
    NAME = params["Name"]
    s = SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PASWD)
    s.sendmail(GMAIL_ID, email, f"Subject: Birthday\n\n{msg} {name}\nFrom {NAME}")
    s.quit()
    print(f"Message:{msg} {name}\nsent to {email} ")


if __name__ == '__main__':
    today = datetime.now().strftime("%d-%m")
    this_year = datetime.now().strftime("%Y")
    df = read_excel(r"birthday.xlsx")
    indexes = []
    try:
        for index, item in df.iterrows():
            if item['Date'] == today and this_year not in str(item['Year']):
                send_email(item['Name'], item['Email'], item['Dialogue'])
                indexes.append(index)
    except Exception as e:
        pass
    for i in indexes:
        df.loc[i, 'Year'] = str(df.loc[i, 'Year']) + ", " + this_year
        df.to_excel('birthday.xlsx', index=False)
