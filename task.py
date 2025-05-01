import os
from RPA.PDF import PDF
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Browser.Selenium import Selenium


class Task:
    """First level of Web Automation Certificates.

    The python class open the website and login. After login, download the excel file and fill the register form Auto,
    take screenshot and create pdf of Sales Data.
    """
    def __init__(self):
        self.pdf = PDF()
        self.http = HTTP()
        self.file = Files()
        self.browser = Selenium()

    def open_website(self):
        """open website"""
        self.browser.open_available_browser("https://robotsparebinindustries.com/")

    def log_in_webssite(self):
        """login website"""
        self.browser.wait_until_element_is_visible('//input[@id="username"]',timeout=10)
        self.browser.input_text('//input[@id="username"]',"maria")
        self.browser.input_text('//input[@id="password"]',"thoushallnotpass")
        self.browser.click_button('//button[@type="submit"]')

    def fill_and_submit_sales_form(self,data):
        """fill register form"""
        self.browser.wait_until_element_is_visible('//input[@id="firstname"]')
        self.browser.input_text('//input[@id="firstname"]', text=data['First Name'])
        self.browser.input_text('//input[@id="lastname"]', text=data['Last Name'])
        self.browser.select_from_list_by_value('//select[@id="salestarget"]',f'{data["Sales Target"]}')
        self.browser.input_text('//input[@id="salesresult"]', text=data['Sales'])
        self.browser.click_button('//button[@type="submit"]')

    def download_excel_file(self):
        """download excel file"""
        excel_file_path = os.path.join(os.getcwd(), 'output/SalesData.xlsx')
        self.http.download(url='https://robotsparebinindustries.com/SalesData.xlsx',target_file=excel_file_path,  overwrite=True)

    def read_excel_and_fill_form(self):
        """read excel file and fill register form"""
        excel_file_path = os.path.join(os.getcwd(), 'output/SalesData.xlsx')
        self.file.open_workbook(excel_file_path)
        excel_data = self.file.read_worksheet_as_table("data", header=True)
        self.file.close_workbook()
        for row in excel_data:
            self.fill_and_submit_sales_form(row)

    def take_sales_summary_screenshot(self):
        """take sales summary screenshot"""
        excel_file_path = os.path.join(os.getcwd(), 'output/sales_summary.png')
        self.browser.capture_element_screenshot('//div[@role="alert"]', filename=excel_file_path)

    def create_pdf_sales_result(self):
        """create pdf of sales result"""
        self.browser.wait_until_element_is_visible('//div[@id="sales-results"]',timeout=10)
        html_data = self.browser.get_element_attribute('//div[@id="sales-results"]','innerHTML')
        pdf_path = os.path.join(os.getcwd(), 'output/Sales_Result.pdf')
        self.pdf.html_to_pdf(html_data, pdf_path)

    def log_out_website(self):
        """logout website"""
        self.browser.wait_until_element_is_visible('//button[@id="logout"]',timeout=10)
        self.browser.click_button('//button[@id="logout"]')


task = Task()
task.open_website()
task.log_in_webssite()
task.download_excel_file()
task.read_excel_and_fill_form()
task.take_sales_summary_screenshot()
task.create_pdf_sales_result()
task.log_out_website()