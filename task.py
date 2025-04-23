from RPA.Browser.Selenium import Selenium

class Task:
    def __init__(self):
        self.browser = Selenium()

    def open_website(self):
        self.browser.open_available_browser("https://robotsparebinindustries.com/")



task = Task()
task.open_website()
