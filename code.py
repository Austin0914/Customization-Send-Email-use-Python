import smtplib
import pandas as pd
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from pathlib import Path
from string import Template

#讀取資料
df = pd.read_excel(io.BytesIO(uploaded['test.xlsx']))
arr = df[df['mail'] != " " ]
arr = arr.fillna("")
print(arr)

receiverList = [] #防止重複送信的陣列
try:
    
    
    # 設定smtp伺服器
    server = smtplib.SMTP(host="smtp.gmail.com", port="587")
    server.ehlo()                               # 驗證SMTP伺服器
    server.starttls()                             # 建立加密傳輸
    server.login("clcp-team@gs.clhs.tyc.edu.tw", "bnykmqdznlkvqnhx")     # gmail帳號/應用程式密碼

    # 讀取信件html
    template = Template(Path("mail.html").read_text())

    # 寄送郵件
    for index, row in arr.iterrows():
        getReceiveCols = row['Name'], row['mail'], row['account'], row['password']   # 讀取相關訊息
        content = MIMEMultipart()                             # 建立MIMEMultipart物件
        content["subject"] = "【壢中資培第三屆成員考】活動提醒及須知"           # 郵件標題
        content["from"] = "clcp-team@gs.clhs.tyc.edu.tw"                  # 寄件者
        content["to"] = getReceiveCols[1]                         #收件者
      
        if getReceiveCols[1] in receiverList:                       # 防止重複寄信
            print("The receiver duplicated")
        else:  
            body = template.substitute({ "name": getReceiveCols[0] ,"account":getReceiveCols[2],"password":getReceiveCols[3]}) # 客製化信件內容的參數
            content.attach(MIMEText(body, "html"))
            server.send_message(content)
            receiverList.append(getReceiveCols[1])
            print("Sent to " + getReceiveCols[1])
    

except Exception as e:
    print("Exception: ", e)
finally:
    server.quit()
    print("Finished")
