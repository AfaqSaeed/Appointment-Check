import requests
import bs4 as beautifulsoup
import time
import os 
import win32com.client
ol=win32com.client.Dispatch("outlook.application")
olmailitem=0x0 #size of the new email
newmail=ol.CreateItem(olmailitem)
newmail.Subject= 'Appointment Mail'
newmail.To='afaqsaeed60@gmail.com'
newmail.CC='ahmmadaniq@gmail.com'
newmail.Body= 'Appointment Openened Contact Afaq.'
attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
newmail.Attachments.Add(attach)
# To display the mail before sending it
newmail.Display() 


url = "https://service2.diplo.de/rktermin/extern/appointment_showForm.do?locationCode=isla&realmId=108&categoryId=1600" ###### Embassy Url 
url = "https://afaqsaeed.github.io/checkweb.html" ########## Test Url 

headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
sent = False
while(True):
    data = requests.get(url,headers=headers).text
    current = ""

    soup = beautifulsoup.BeautifulSoup(data,'html.parser')
    all = soup.findAll("select")[-1]
    
    if not  os.path.exists("data.html"):
        with open("data.html","w+") as f:
            f.write(str(all))

    with open("data.html","r") as f:
        backup = f.read()
        
    if str(all) == backup:
        print("same")

    else:
        if not sent:
            newmail.Send()
            sent=True
        print("changed")
        

    time.sleep(5)


