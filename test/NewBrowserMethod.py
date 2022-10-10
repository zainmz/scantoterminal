from selenium import webdriver
from selenium.webdriver.chrome.options import Options

opt = Options()
opt.add_experimental_option("debuggerAddress", "localhost:9988")
driver = webdriver.Chrome(executable_path=r"D:\Python Learning\scanToTerminal\chromedriver.exe",chrome_options=opt)
driver.get('http://facebook.com')
