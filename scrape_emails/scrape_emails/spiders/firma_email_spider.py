import scrapy
import openpyxl as openpyxl
from scrapy import signals
from scrapy.signalmanager import dispatcher
import math


class QuotesSpider(scrapy.Spider):
    name = "firms"
    accumulated_items = []
    all_page_urls = []
    items_per_page = 25
    # you have to provide the filename
    EXCEL_FILE_NAME = "output.xlsx"
    # you have to provide the page url
    BASE_LINK = "https://www.zoznam.sk/katalog/Spravodajstvo-informacie/Abecedny-zoznam-firiem/0-9/sekcia.fcgi?sid=1172&so=&page={}&desc=&shops=&kraj=&okres=&cast=&attr="
    start_urls = [BASE_LINK.format(1)]

    def __init__(self, *args, **kwargs):
        super(QuotesSpider, self).__init__(*args, **kwargs)
        # Connect the spider_closed signal to the handler
        dispatcher.connect(self.handle_spider_closed, signal=signals.spider_closed)

    def parse(self, response):
        firma = response.css("small::text")[0].extract()
        formatted_page_number = int(firma.strip("()"))
        last_page = math.ceil(formatted_page_number / self.items_per_page)
        for i in range(1, last_page + 1):
            yield scrapy.Request(url=self.BASE_LINK.format(i), callback=self.parse_firma)

    def parse_firma(self, response):
        BASE_FIRMA_LINK = "https://www.zoznam.sk"  # base link for profiles (do not modify)
        firma = response.css("a.link_title::attr(href)").extract()
        for link in firma:
            yield scrapy.Request(url=f'{BASE_FIRMA_LINK}{link}', callback=self.parse_email)

    def parse_email(self, response):
        emails = response.css('.profile .row .col-sm-9 a::text').extract()
        for email in emails:
            if "@" in email:
                self.accumulated_items.append([email])
                break

    def handle_spider_closed(self, spider, reason):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # write the Excel file
        for row_data in self.accumulated_items:
            sheet.append([str(cell) for cell in row_data])
        workbook.save(self.EXCEL_FILE_NAME)
