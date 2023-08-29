import scrapy
import openpyxl as openpyxl
from scrapy import signals
from scrapy.signalmanager import dispatcher
import math


class QuotesSpider(scrapy.Spider):
    # CONFIG ITEMS
    items_per_page = 25  # you can modify items per page if needed
    first_page = 1  # you can modify if you don't want to start from the first page
    last_page = False  # you can provide the last page if you don't want all pages
    # you have to provide the filename
    EXCEL_FILE_NAME = "output.xlsx"  # you can modify the output file
    # you have to provide the page url, IMPORTANT! - put '{}' after '&page=' because it's a dynamic link
    BASE_LINK = "https://www.zoznam.sk/katalog/Spravodajstvo-informacie/Abecedny-zoznam-firiem/0-9/sekcia.fcgi?sid=1172&so=&page={}&desc=&shops=&kraj=&okres=&cast=&attr="

    # !!!!! IMPORTANT
    name = "firms"  # you can run the scraping with 'scrapy crawl firms' command

    # DO NOT MODIFY THESE ITEMS
    accumulated_emails = []  # container for emails
    all_page_urls = []  # container for urls
    BASE_FIRMA_LINK = "https://www.zoznam.sk"  # base link for profiles (do not modify)
    start_urls = [BASE_LINK.format(1)]  # first page for scrape all page numbers

    def __init__(self, *args, **kwargs):
        super(QuotesSpider, self).__init__(*args, **kwargs)
        # Connect the spider_closed signal to the handler
        dispatcher.connect(self.handle_spider_closed, signal=signals.spider_closed)

    def parse(self, response):
        if not self.last_page:  # extract the last page element if the last page not provided
            firma = response.css("small::text")[0].extract()  # extract items number
            formatted_page_number = int(firma.strip("()"))  # convert the number to int
            # subtract the items with the items per page
            self.last_page = math.ceil(formatted_page_number / self.items_per_page)
        # loop through all pages or between the provided page numbers
        for i in range(self.first_page, self.last_page + 1):
            # send the page to the parse firma method
            yield scrapy.Request(url=self.BASE_LINK.format(i), callback=self.parse_firma)

    def parse_firma(self, response):  # extract the page links from the page
        firma = response.css("a.link_title::attr(href)").extract()
        for link in firma:  # send the links to the parse email method
            yield scrapy.Request(url=f'{self.BASE_FIRMA_LINK}{link}', callback=self.parse_email)

    def parse_email(self, response):  # extract emails from the users page
        emails = response.css('.profile .row .col-sm-9 a::text').extract()
        for email in emails:
            if "@" in email:
                self.accumulated_emails.append([email])  # add emails to the container
                break

    def handle_spider_closed(self, spider, reason):  # handle close
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # write the Excel file
        for row_data in self.accumulated_emails:
            sheet.append([str(cell) for cell in row_data])
        workbook.save(self.EXCEL_FILE_NAME)
