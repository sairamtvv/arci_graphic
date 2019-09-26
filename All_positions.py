#Large interdelay which shall be in days usually
       
#        largevar = StringVar(value=self.largedelaydefault)
#        self.largeentry=TK.Entry(textvariable=largevar)
#        self.largeentry.grid(row=5,column=1)
         
        #label for largedelay
#        self.labellargeentry=TK.Label(text="LARGE DELAY (mins)", fg='red')
#        self.labellargeentry.grid(row=4,column=1)
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import sys
#import http.client
#import pprint
#
#connection = http.client.HTTPSConnection("file:///D:/day2/allitems/friday_data/WebServerPort_again.json")
#connection.request("GET", "/")
#response = connection.getresponse()
#headers = response.getheaders()
#pp = pprint.PrettyPrinter(indent=4)
#pp.pprint("Headers: {}".format(headers))




driver=webdriver.Chrome("C:\\Users\\sairamtvv\\Videos\\chromedriver_win32\\chromedriver.exe")

#Get the html of the site, basically ip address here.

baseUrl = "file:///D:/day2/allitems/friday_data/After_typing_rapid_X_0.54200.htm"

driver.get(baseUrl)
time.sleep(0.5)


#Clicking two balls on the screen
balls_Element=driver.find_element_by_id("enableDisableAxis0")

if balls_Element is None:
    print("The position controller is not connecting...")
    sys.exit()
else:
    balls_Element.click()
    time.sleep(1)
    driver.implicitly_wait(1)


#Clicking home button
driver.find_element_by_id("homeAxis0").click()
driver.implicitly_wait(1)



#sending ENABLE X and pressing enter
imme_comm=driver.find_element_by_id("immediate-command-text")
imme_comm.send_keys("ENABLE X")
time.sleep(1)
imme_comm.send_keys(Keys.RETURN)
time.sleep(0.2)


#checking for NO ERROR in the  bottom status bar
#bottombar_Element=driver.find_element_by_id("status-bar")
#if bottombar_Element is not None:
#    print("bottombar_Element found")
#bottombar_value=bottombar_Element.get_attribute("value")
#print("The bottombar value is " +bottombar_value)


check_enable_Element = driver.find_element_by_id('axis0Status')
check_enable_Text = check_enable_Element.text

if check_enable_Text=='Enabled':
    print("Enabled and lets continue")
    time.sleep(0.5)

#clearing the immediate-command text for next command and then Absolute command
imme_comm=driver.find_element_by_id("immediate-command-text")
time.sleep(0.5)
imme_comm.clear()
imme_comm.send_keys("ABSOLUTE")
time.sleep(1)
imme_comm.send_keys(Keys.RETURN)
time.sleep(0.2)


#clearing the immediate-command text for next command and then POSITION 1
imme_comm=driver.find_element_by_id("immediate-command-text")
time.sleep(0.5)
imme_comm.clear()
imme_comm.send_keys("RAPID X -0.542000 F5")
time.sleep(1)
imme_comm.send_keys(Keys.RETURN)
time.sleep(0.2)
driver.implicitly_wait(1)

#Checking if the position feedback has reached the desired value
pos_feedback_Text=0
desired_pos=-0.542000
while abs(desired_pos-float(pos_feedback_Text))>10**-3:
    pos_feedback_Element=driver.find_element_by_id('axis0PosFbk')
    time.sleep(0.5)
    pos_feedback_Text = pos_feedback_Element.text
print("Reached the desired position")


print("First test completed")