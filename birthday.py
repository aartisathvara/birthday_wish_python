import pandas as pd
import datetime
import smtplib
import os

os.chdir(r"G:\PycharmProjects\Basic Python\birthday")
os.mkdir("testing")
#enter your gmail id hand password
GMAIL_ID='aarusathvara@gmail.com'
GMAIL_PASSWORD='Aarti@$3108'


def send_email(to,sub,msg):
    print(f"Email to {to} sent with subject: {sub} and Message {msg}")
    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PASSWORD)
    s.sendmail(GMAIL_ID,to,f"subject:{sub}\n\n{msg}")
    s.quit()


if __name__=="__main__":
    # send_email(GMAIL_ID,"subject"," Test message")

    #data excel file read
    df=pd.read_excel(r"G:\PycharmProjects\Basic Python\birthday\Data.xlsx")
    # print(df)

    # df2= pd.read_csv(r"G:PycharmProjects\Basic Python\birthday\Data.csv")
    # print("df2",df2)

    #now get today date
    today=datetime.datetime.now().strftime("%d-%m")
    yearNow=datetime.datetime.now().strftime("%Y")
    # print(today)

    writeInd=[]
    #now excel file iterrate
    for index,item in df.iterrows():
        # print(index,item['Birthdate'])

        bday=item['Birthdate'].strftime("%d-%m")
        # print(bday)

        if (today==bday) and yearNow not in str(item['Year']) :
            send_email(item['Email'], "Happy Birthday", item['Dialogue'])
            writeInd.append(index)

    print(writeInd)
    for i in writeInd:
        yr=df.loc[i,'Year']
        df.loc[i,'Year']=str(yr)+','+str(yearNow)
        print(df.loc[i,'Year'])
    # print(df)
    df.to_excel('Data.xlsx')
