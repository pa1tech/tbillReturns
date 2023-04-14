import re, time,os
import pandas as pd
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

driver = Chrome(service=Service(ChromeDriverManager().install()))
excel_file = "Tbill_Returns_2023.xlsx"

try:
	df = pd.read_excel(excel_file,sheet_name="Sheet1")
	pressIDs = df["Press Release ID"].to_numpy(dtype=int)
except:
	pressIDs = []
	df = pd.DataFrame()


# Open website and select all press releases of 2023
driver.get("https://www.rbi.org.in/Scripts/BS_PressReleaseDisplay.aspx")
if(re.search("Your support ID is:*",driver.page_source)): input("\n Press enter to cotinue after completing captcha:")

driver.find_element(By.ID,'btn2023').click()
if(re.search("Your support ID is:*",driver.page_source)): input("\n Press enter to cotinue after completing captcha:")

driver.find_element(By.ID,'20230').click()
if(re.search("Your support ID is:*",driver.page_source)): input("\n Press enter to cotinue after completing captcha:")

text = driver.page_source
table = re.search('<table class="tablebg".*</table>',text)
if(table):
	tableText = text[table.start():table.end()]
tableText = tableText.split("</tr>")


for text in tableText:
	prid = re.findall('[0-9]+">91 days, 182 days and 364 days T-Bill Auction Result: Cut off',text)
	if(len(prid)>0):
		prid = prid[0].split('"')[0]
		if(int(prid) not in pressIDs):
			driver.get("https://www.rbi.org.in/Scripts/BS_PressReleaseDisplay.aspx?prid="+prid)
			if(re.search("Your support ID is:*",driver.page_source)): input("\n Press enter to cotinue after completing captcha:")

			try:
				with open("tmp.html", "w", encoding="utf-8") as f:
					f.write(driver.page_source)

				dfs = pd.read_html("tmp.html")

				d = dfs[2]
				returns = [d.loc[2,2],d.loc[2,3],d.loc[2,4]]
				for i in range(len(returns)):
					try:
						returns[i] = ((returns[i].split(":"))[1])[:-1]
					except:
						returns[i] = "NA"

				d = dfs[0]
				date = d.loc[1,0].split(":")[1]

				df1 = pd.DataFrame({"Press Release ID":[prid], "Date":[date], "91 DTB Yield":[returns[0]], "182 DTB Yield":[returns[1]], "364 DTB Yield":[returns[2]]})
				df = pd.concat((df,df1))
			except :
				print("Exception in https://www.rbi.org.in/Scripts/BS_PressReleaseDisplay.aspx?prid="+prid)

driver.quit()

df['Date'] = pd.to_datetime(df['Date'])
df = df.sort_values(by='Date',ascending=True, ignore_index=True)
df['Date'] = df['Date'].dt.strftime('%d %b, %Y')
df.to_excel(excel_file,sheet_name="Sheet1",index = False)

os.remove("tmp.html")