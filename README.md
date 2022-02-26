# Customization-Send-Email-use-Python
# Python>Gmail客製化信件大量發送
> [name=Austin]
## 誰要用SMTP寄信
1. 需要大量客製化信件的人(像是每個信件有同的帳號密碼)
2. 沒有Outlook的人
3. 被這東西搞到發瘋的人

## SMTP

簡單郵件傳輸協定 (Simple Mail Transfer Protocol, SMTP) 是在Internet傳輸email的標準。

### 發信協定: SMTP [LINK](https://wanteasy.com.tw/doc/imap-pop3-smtp-difference.html)

相較於IMAP和POP3是在討論如何把郵件收到電腦上，但我們總需要把信寄出去吧? 而這發信、寄信的功能就是SMTP，SMTP的全名是Simple Mail Transfer Protocol(簡單郵件傳輸協定)。我們在Outlook或Thunderbird上設定SMTP主機和登入資訊，就是設定當我們要寄出郵件時，會透過這個SMTP主機幫我們把郵件傳送到收件者的郵件主機。

---

## 取得Google 應用程式密碼
```diff
@@ 用完之後請迅速刪除密碼，有被盜的風險!!!!@@
@@ 用完之後請迅速刪除密碼，有被盜的風險!!!!@@
@@ 用完之後請迅速刪除密碼，有被盜的風險!!!!@@
```
1. 管理你的google帳號

 ![](https://i.imgur.com/0T2JxxX.png)
 
2. 安全性>雙步驗證(開啟)

![](https://i.imgur.com/bDYPRnu.png)

3. 應用程式密碼>產生

![](https://i.imgur.com/v4s27uc.png)

4. 取得密碼(圖片密碼無效!)

![](https://i.imgur.com/3y15VsB.png)


---

## 使用原理

### 讀取資料>SMTP伺服器設定>html內文撰寫>個別送出同時替換相關參數
![](https://i.imgur.com/Jn9Vdwp.png)

## 使用方法

### 讀取excel
```python
df = pd.read_excel(io.BytesIO(uploaded['test.xlsx']))
arr = df[df['mail'] != " " ]
arr = arr.fillna("")
print(arr)
```
L1:讀取excel資料
L2:如果mail不是空的就讀去資料
備註::讀取資料為google colab寫法
### 設定SMTP伺服器
```python
server = smtplib.SMTP(host="smtp.gmail.com", port="587")
server.ehlo()                               # 驗證SMTP伺服器
server.starttls()                             # 建立加密傳輸
server.login("clcp-team@gs.clhs.tyc.edu.tw", "bnykmqdznlkvqnhx") 
```
L1、2、3:伺服器設定(固定)
L4:寄送者信箱/google應用程式密碼
### html 內文撰寫
因為是信件head不用寫
[HTML簡易語法](https://tw.alphacamp.co/blog/html-guide)
* 標題：h1, h2, h3, h4, h5, h6
* 文字段落：p
* 清單：ul, ol, li
* 強調語氣：em, strong
* 換行：br
* 水平線：hr
* 超連結：a
* 圖片：img
* 區域：div, span
* 表格：table, tr, th, td
* 表單：form, label, input

`要替換的參數前面放 $變數`

```html
<!DOCTYPE html>
<html lang="en">
    <head>
    </head>
    <body>
            親愛的<strong>$name</strong> 你的帳號密碼為<br><p>
        
            <h3>帳號密碼</h3>
            
            帳號: $account<br>
            密碼: $password
    </body>
</html>
```
### 讀取html
```python
template = Template(Path("mail.html").read_text())
```
### 個別寄信
```python
for index, row in arr.iterrows():
        getReceiveCols = row['Name'], row['mail'], row['account'], row['password']   
        content = MIMEMultipart()                             
        content["subject"] = "【壢中資培第三屆成員考】活動提醒及須知"           
        content["from"] = "clcp-team@gs.clhs.tyc.edu.tw"                  
        content["to"] = getReceiveCols[1]                         
      
        if getReceiveCols[1] in receiverList:                     
            print("The receiver duplicated")
        else:  
            body = template.substitute({ "name": getReceiveCols[0] ,"account":getReceiveCols[2],"password":getReceiveCols[3]}) 
            content.attach(MIMEText(body, "html"))
            server.send_message(content)
            receiverList.append(getReceiveCols[1])
            print("Sent to " + getReceiveCols[1])
```
L1:執行n次寄送信件
L2:讀取excel檔第一行資訊(標籤)
L3:建立MIMEMultipart物件
L4:信件標題
L5:寄件者信箱
L6:收件者gmail
L11:替換相關參數
L13:寄送


## CODE
> [google colab(可以直接編譯//學校帳號不能用)](https://colab.research.google.com/drive/1DpPhpaEQ7hfTO_TQOH8W0CA8RkvStF1O?usp=sharing)
## 參考文章
 [Python實現Gmail客製化信件大量發送 — 【基礎篇】](https://marketingliveincode.com/?p=185)
 
 [用 Python 寄送客制化郵件 ( Email )](https://ycjhuo.gitlab.io/blogs/Python-Mutiple-And-Customize-Mail-Message.html#%E5%AF%84%E5%87%BA%E9%83%B5%E4%BB%B6)
 
[Python實戰應用_Python寄送Gmail電子郵件實作教學](https://www.learncodewithmike.com/2020/02/python-email.html)
