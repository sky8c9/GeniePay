import os, shutil, sys, holidays
import pandas as pd
import numpy as np
from pytz import timezone
from threading import Thread
from time import sleep
from datetime import date, datetime, timedelta
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

CHROME_DRIVER_PATH = "replace with path to chromedriver.exe"
DATA_PATH = "replace with path to template_data.xlsx"
PAYMENT_RECEIPT_PATH = "replace with path to receipt folder"
PAYMENT_SUMMARY_PATH = "repace with path to payment history folder"

URL = "https://www.eftps.gov/"
user_agent = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36"

class PayBot(Thread):
	def __init__(self, opt, data):
		super(PayBot, self).__init__()

		chrome_options = Options()
		chrome_options.add_argument(f'user-agent={user_agent}')
		chrome_options.add_argument("--headless")
		chrome_options.add_argument("--window-size=2560,1440")
		chrome_options.add_argument('--ignore-certificate-errors')
		chrome_options.add_argument('--allow-running-insecure-content')
		chrome_options.add_argument("--disable-extensions")
		chrome_options.add_argument("--proxy-server='direct://'")
		chrome_options.add_argument("--proxy-bypass-list=*")
		chrome_options.add_argument("--start-maximized")
		chrome_options.add_argument('--disable-gpu')
		chrome_options.add_argument('--disable-dev-shm-usage')
		chrome_options.add_argument('--no-sandbox')

		ser = Service(CHROME_DRIVER_PATH)
		self.driver = webdriver.Chrome(service=ser, options=chrome_options)

		# unzip login data fields
		self.name, self.ein1, self.ein2, self.pin, self.password, self.form, self.quarter, self.year, self.amount = data
		
		# Reformat ein2 & pin
		self.ein2 = str(self.ein2).zfill(7)
		self.pin = str(self.pin).zfill(4)

		# other fields
		self.type = opt
		self.finished = False
		self.tpaid = 0.0

	def gen_error(self, cause):
		print(self.name + " " + cause)
		sys.exit(1)

	def run(self):
		print(self.name + "...")

		if (self.type == 2):
			self.pay()
		elif (self.type == 3):
			self.screenshot_payment_history()
		else:
			self.total_pay()

		self.finished = True
	
	def pay(self):
		self.log_in()
		self.select_deposit()
		self.set_up_payment()
		self.log_out()

	def log_in(self):
		try:
			# Load Url
			self.driver.get(URL)
			self.driver.find_element(By.LINK_TEXT, "MAKE A PAYMENT").click()
			sleep(5)

			# Login
			self.driver.find_element(By.NAME, "EIN1").send_keys(str(self.ein1))
			self.driver.find_element(By.NAME, "EIN2").send_keys(str(self.ein2))
			self.driver.find_element(By.NAME, "PIN").send_keys(str(self.pin))
			self.driver.find_element(By.NAME, "password").send_keys(self.password)
			self.driver.find_element(By.XPATH, "//input[@value='Login']").click()
		except:
			self.gen_error("Loading Login Page Error!!!")

	def log_out(self):
		try:
			self.driver.find_element(By.LINK_TEXT, "LOGOUT").click()
			self.driver.find_element(By.LINK_TEXT, "ACCEPT").click()
			self.driver.quit()
		except:
			self.gen_error("Logout Error!!!")

	def select_deposit(self):
		try:
			# Select Tax Form
			form_select = Select(self.driver.find_element(By.NAME, "commonSelection"))
			form_select.select_by_value(str(self.form))
			self.driver.find_element(By.XPATH, "//input[@value='next']").click() 

			# Select Federal Tax Deposit
			self.driver.find_element(By.NAME, "taxType").click()
			self.driver.find_element(By.XPATH, "//input[@value='next']").click() 
		except:
			self.gen_error("Login Error!!!")

	def set_up_payment(self):
		try:
			# Clear form
			self.driver.find_element(By.LINK_TEXT, "CLEAR FORM").click()

			# Enter amount
			self.driver.find_element(By.ID, "singlePayment.amount.value").send_keys(str(format(self.amount, '.2f')))

			# Select quarter
			quarter_select = Select(self.driver.find_element(By.NAME, 'singlePayment.taxPeriodMonth'))
			quarter_select.select_by_value(str(self.quarter * 3))

			# Enter year
			self.driver.find_element(By.NAME, "singlePayment.taxPeriodYear").send_keys(str(self.year))

			# Select payment date within business day
			paid_day = date.today()
			hour = datetime.now(timezone('EST')).hour
			if (hour >= 14):
				paid_day += timedelta(1)

			us_holidays = holidays.UnitedStates()
			while ((paid_day.weekday() >= 5) or (paid_day in us_holidays)):
				paid_day += timedelta(1)

			self.driver.find_element(By.NAME, "singlePayment.settlementDate.dateString").send_keys(paid_day.strftime("%m/%d/%Y"))
			self.driver.find_element(By.XPATH, "//input[@value='next']").click() 
			self.driver.find_element(By.XPATH, "//input[@value='next']").click()

			# Confirm payment
			self.driver.find_element(By.XPATH, "//input[@value='makePayment']").click()

			# Receipt screenshot
			fname = self.name + date.today().strftime("%Y-%m-%d") + "-receipt.png"
			self.driver.save_screenshot(PAYMENT_RECEIPT_PATH + "/" + fname)
		except:
			self.gen_error("Setup Payment Error!!!")

	def load_payment_history(self):
		# 180days payment view
		self.log_in()
		try:
			self.driver.find_element(By.LINK_TEXT, "CHECK PAYMENT HISTORY").click()
			self.driver.find_element(By.ID, "180days").click()
			self.driver.find_element(By.XPATH, "//input[@title='Search']").click() 
		except:
			self.gen_error("Load Payment History Error!!!")

	def screenshot_payment_history(self):
		self.load_payment_history()
		fname = self.name + date.today().strftime("%Y-%m-%d") + ".png"
		self.driver.save_screenshot(PAYMENT_SUMMARY_PATH + "/" + fname)
		self.log_out()

	def same_date_payment(self):
		self.load_payment_history()
		most_recent_pay_date = payments[0].text.split()[1]
		return most_recent_pay_date == date.today().strftime("%Y-%m-%d") 

	def total_pay(self):
		self.load_payment_history()

		try:
			payments = self.driver.find_elements(By.XPATH, "//table[@summary='Payment records']//*[@align='center']")

			# Compute total pay for a current quarter
			for pay in payments:
				setDate, initDate, taxForm, taxPeriod, paid_amount, status = pay.text.split()

				if (taxForm == str(self.form)):
					year, quarter = taxPeriod.split("/")

					if (int(quarter[-1]) == self.quarter and int(year) == self.year):
						self.tpaid += float(paid_amount.replace(',', '').replace('$', ''))	     
		except:
			self.gen_error("Total Pay Error!!!")

		self.log_out()       

def clean_up(fpath):
	for files in os.listdir(fpath):
		path = os.path.join(fpath, files)
		try:
			shutil.rmtree(path)
		except OSError:
			os.remove(path)

def color_schema(df, states):
	tmp = df.copy()[states]
	tmp.loc[:] = 'background-color: red'  
	return tmp

def check_EIN(df):
	dup = df.duplicated(subset=['EIN1', 'EIN2']).to_numpy()
	if (len(df[dup]) > 0):
		print("Please remove duplicate EIN entries:")
		print(df[dup])
		sys.exit(1)

def menu():
	print("-----------------------------")
	print("1: View Data")
	print("2: Make Payment")
	print("3: Screenshot Payment History")
	print("4: Compute Total Pay")
	print("-----------------------------")

def task():
	option = int(input("Enter: "))
	df = pd.read_excel(DATA_PATH)
	check_EIN(df)

	if (option == 1):
		print(df.head())
	elif (option == 2 or option == 3 or option == 4):
		# Confirm payment - safety feature
		if (option == 2):
			confirm = input("Type 'Yes' to confirm payment: ");

			if (confirm != 'Yes'):
				print("No Payment - Exit!!!")
				return

			clean_up(PAYMENT_RECEIPT_PATH)

		if (option == 3):
			clean_up(PAYMENT_SUMMARY_PATH)

		# Spawn workers
		threads = []
		for index, row in df.iterrows():
			worker = PayBot(option, row)
			threads.append(worker)
			worker.start()

		for thread in threads:
			thread.join()

		# Failed payments shown in red
		if (option == 2):
			states = [not o.finished for o in threads]
			out = df.style.apply(lambda f : color_schema(df, states), axis=None)
			out.to_excel('Payment_Status.xlsx', engine='openpyxl')

		# Total paid summary
		if (option == 4):
			df['Tax Period'] = "Q" + df['Quarter'].astype(str) + "/" + df['Year'].astype(str)
			df['Total'] = [o.tpaid for o in threads]
			df[['Tax Period', 'Name', 'Total']].to_excel('Paid_Summary.xlsx', engine='openpyxl')

def main():
	menu()
	task()
	
if __name__ == "__main__":
	main()
