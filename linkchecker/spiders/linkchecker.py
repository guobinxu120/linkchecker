# -*- coding: utf-8 -*-
from scrapy import Spider, Request
from collections import OrderedDict
from xlrd import open_workbook

def readExcel(path):
    wb = open_workbook(path)
    result = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        herders = []
        for row in range(0, number_of_rows):
            values = OrderedDict()
            for col in range(number_of_columns):
                value = (sheet.cell(row,col).value)
                if row == 0:
                    herders.append(value)
                else:

                    values[herders[col]] = value
            if len(values.values()) > 0:
                result.append(values)
        break

    return result


class AngelSpider(Spider):
    name = "linkchecker"
    start_urls = 'https://www.amazon.fr/'
    count = 0
    use_selenium = False
    models = readExcel("Links checker.xlsx")

    def start_requests(self):
        for i, val in enumerate(self.models):
            url = val['URL LINKS']
            yield Request(url , callback=self.parse1, errback=self.parse, meta={'order_num':i}, dont_filter=True)

    def parse1(self, response):
        order_num = response.meta['order_num']
        self.models[order_num]['STATUTS'] = 200
        yield self.models[order_num]

    def parse(self, response):
        order_num = response.request.meta['order_num']
        self.models[order_num]['STATUTS'] = 404
        yield self.models[order_num]
