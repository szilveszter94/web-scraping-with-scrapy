import scrapy
import openpyxl as openpyxl
from scrapy import signals
from scrapy.signalmanager import dispatcher


class ZlatestrankySpider(scrapy.Spider):
    # CONFIG ITEMS
    first_page = 1  # you can modify items per page if needed
    max_page = 12908  # you can modify if you don't want to load all pages MAX PAGE: 12908
    BASE_LINK = "https://www.zlatestranky.sk/firmy/-/q_*/{}/"
    # you have to provide the filename
    EXCEL_FILE_NAME = "output.xlsx"  # you can modify the output file
    # you have to provide the page url, IMPORTANT! - put '{}' after '&page=' because it's a dynamic link

    # !!!!! IMPORTANT
    name = "zlatestranky"  # you can run the scraping with 'scrapy crawl zlatestranky' command

    # DO NOT MODIFY THESE ITEMS
    accumulated_emails = []  # container for emails
    start_urls = []
    for i in range(1, max_page):
        start_urls.append(BASE_LINK.format(i))

    def __init__(self, *args, **kwargs):
        super(ZlatestrankySpider, self).__init__(*args, **kwargs)
        # Connect the spider_closed signal to the handler
        dispatcher.connect(self.handle_spider_closed, signal=signals.spider_closed)

    def parse(self, response):
        emails = response.css(".mail a::text").extract()
        for email in emails:
            if "@" in email:
                self.accumulated_emails.append([email])

    def handle_spider_closed(self, spider, reason):  # handle close
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # write the Excel file
        for row_data in self.accumulated_emails:
            sheet.append([str(cell) for cell in row_data])
        workbook.save(self.EXCEL_FILE_NAME)