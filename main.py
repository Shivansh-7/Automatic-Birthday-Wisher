import pandas as pd 
import datetime
import smtplib
import os
os.chdir(r"C:\Users\ucc14\PycharmProjects\Birthday Wisher")
os.mkdir("testing")

GMAIL_ID = ''
GMAIL_PASS = ''
def sendEmail(to, sub, msg):
    print(f"Email to {to} sent with subject: {sub} and message {msg}")
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PASS)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()

if __name__ == "__main__":
    # sendEmail(GMAIL_ID, "subject", "test-message")
    # exit()
    df = pd.read_excel("Birthday-data-file.xlsx")
    # print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime(("%Y"))
    # print("yearNow",yearNow)

    writeInd = []
    for index, item in df.iterrows():
        # print(index, item['Birthday'])
        bday = item['Birthday'].strftime("%d-%m")
        print(bday)
        if (today == bday) and yearNow not in str(item['Year']):
            sendEmail(item['Email'], "Happy Birthday", item['Dialogue'])
            writeInd.append(index)

    print(writeInd)
    if writeInd == []:
        print("No friends is left are wishing his birthday")
    else:
        for i in writeInd: 
            yr = df.loc[i, 'Year']
            df.loc[i, 'Year'] = str(yr) + ',' + str(yearNow)
        
        # print(df)
        df.to_excel('Birthday-data-file.xlsx', index=False)

   

