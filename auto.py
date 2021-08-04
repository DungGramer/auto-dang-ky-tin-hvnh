from selenium import webdriver
import seleniumrequests
import time
import xlrd

wb = xlrd.open_workbook('dki.xlsx')
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

rows = sheet.nrows
columns = sheet.ncols

driver = driver = webdriver.Chrome(executable_path=r'./webdriver/chromedriver')
driver.get("http://regist.hvnh.edu.vn/Login")
driver.maximize_window()


def login():
  account = open('account.txt', 'r')
  username = driver.find_element_by_xpath('//*[@id="username"]')
  username.send_keys(account.readline())

  password = driver.find_element_by_xpath('//*[@id="password"]')
  password.send_keys(account.readline())

  submit = driver.find_element_by_xpath('/html/body/form/div/div/div[8]/input')
  submit.click()
  print ("Logged successfully")

def searchHP(id):
  driver.get("http://regist.hvnh.edu.vn/TraCuuHocPhan")
  hp = driver.find_element_by_xpath(
      '/html/body/form/div[2]/div[2]/div[2]/div/div[2]/div[1]/b/select')
  hp.click()

  tenHP = driver.find_element_by_xpath('/html/body/form/div[2]/div[2]/div[2]/div/div[2]/div[1]/b/select/option[2]')
  tenHP.click()

  searchHP = driver.find_element_by_xpath('/html/body/form/div[2]/div[2]/div[2]/div/div[2]/div[1]/input[1]')
  searchHP.send_keys(id)

  find = driver.find_element_by_xpath(
      '/html/body/form/div[2]/div[2]/div[2]/div/div[2]/div[1]/input[2]')
  find.click()

def openNewTab(rows, url):
  for row in range(1, rows):
    driver.execute_script("window.open('"+url+"','_blank');")
    driver.switch_to.window(driver.window_handles[row])
    searchHP(sheet.cell_value(row, 0))
    column = sheet.cell_value(row, 1).replace(",", "|").replace(" ", "")
    driver.execute_script('''
      let re = new RegExp(arguments[0], "g");
      [...document.querySelectorAll('td')].filter(td => td.innerText.match(re)).map(selector => {
        console.log(selector);
        selector.style.backgroundColor = 'tomato';
        try { 
          selector.parentElement.querySelector('input[type="radio"]').checked = true;
          document.forms[1].submit();
        } catch (err) {
          console.log(err);
        }
      })
    ''', column)



login()
openNewTab(rows, 'http://regist.hvnh.edu.vn/TraCuuHocPhan')
