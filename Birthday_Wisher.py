from pandas import read_excel
from datetime import datetime
from smtplib import SMTP


def send_email(name, email, msg):
    with open(r"infomail.txt", "r") as f:
        a = f.read()
        a = a.split()
        a = a.replace("\n")
    GMAIL_ID = a[0]
    GMAIL_PASWD = a[1]
    s = SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PASWD)
    s.sendmail(GMAIL_ID, email, f"Subject: Birthday\n\n{msg} {name}\nFrom Meet Pogul and Family")
    s.quit()
    print(f"Message:{msg} {name}\nsent to {email} ")


if __name__ == '__main__':
    # send_email('Isha', "ishasp281004@gmail.com", 'Happy Birthday to you')
    # exit()
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

# Study smtp and less secure app
# task schedular
# create task
# set trigers
# write genreal
# action -> start program
# select python.exe from python installed folder
# add arguments -> "location of file"
