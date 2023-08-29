import scrapy
import openpyxl as openpyxl
from scrapy import signals
from scrapy.signalmanager import dispatcher


class QuotesSpider(scrapy.Spider):
    name = "firms"
    start_urls = []
    accumulated_items = []
    # you have to provide the filename
    EXCEL_FILE_NAME = "output.xlsx"
    # you have to provide the page url
    BASE_LINK = "https://www.zoznam.sk/katalog/Spravodajstvo-informacie/Abecedny-zoznam-firiem/A/sekcia.fcgi?sid=1173&so=&page={}&desc=&shops=&kraj=&okres=&cast=&attr="
    for i in range(1, 2):  # you have to provide the page numbers + 1
        start_urls.append(BASE_LINK.format(i))

    def __init__(self, *args, **kwargs):
        super(QuotesSpider, self).__init__(*args, **kwargs)
        # Connect the spider_closed signal to the handler
        dispatcher.connect(self.handle_spider_closed, signal=signals.spider_closed)


    def parse(self, response):
        BASE_FIRMA_LINK = "https://www.zoznam.sk"  # base link for profiles (do not modify)
        firma = response.css("a.link_title::attr(href)").extract()
        for link in firma:
            yield scrapy.Request(url=f'{BASE_FIRMA_LINK}{link}', callback=self.parse_firma)

    def parse_firma(self, response):
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