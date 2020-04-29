import re
import scrapy
from scrapy.shell import inspect_response

from module.items import CS_FSC_Item
class CS_FSC_Spider(scrapy.Spider):
    name = "cs_fsc"
    custom_settings = {
        'DUPEFILTER_CLASS': 'scrapy.dupefilters.BaseDupeFilter',
    }
    start_urls = [
        "https://www.ncss.gov.sg/GatewayPages/Social-Service-Organisations/Membership/List-of-NCSS-Members"
    ]

    def parse(self, response):
        search_strings = ["community service", "cs", "family service", "fsc", "social service"]
        found = False
        urls = []
        for sel in response.xpath("//*[@class='innerPage']/div/div/div/table/tr[3]/td/a"):
            #inspect_response(response, self)
            found = False
            for s in search_strings:
                if s in sel.xpath('text()')[0].extract().lower():
                    found = True
                    break
            if found:
                urls.append(sel.xpath('@href')[0].extract().strip())
                yield scrapy.Request(sel.xpath('@href')[0].extract().strip(),
                            callback = self.attempt_contact_pages,
                            cb_kwargs = dict(name = sel.xpath('text()')[0].extract(), url = sel.xpath('@href')[0].extract().strip()))
        print("final")
        print(len(urls))
        print(urls)

            
    def attempt_contact_pages(self, response, name, url):
        search_strings = ["contact", "connect"]
        found = False
        for sel in response.xpath("//*/a"):
            for s in search_strings:
                if len(sel.xpath('text()').extract()) > 0 and s in sel.xpath('text()')[0].extract().lower() and not found:
                    found = True
                    next = sel.xpath('@href')[0].extract()
                    if response.url not in sel.xpath('@href')[0].extract():
                        print(next)
                        next = response.url + "/" + next
                        next = next.replace("https://", "").replace("http://", "")
                        while "//" in next:
                            next = next.replace("//", "/")
                        next = "http://" + next
                    yield scrapy.Request(next,
                                callback = self.parse_target_pages,
                                cb_kwargs = dict(name = name, url = url))
                    break
            if found:
                break
        if not found:
            yield scrapy.Request(url,
                            callback = self.parse_target_pages,
                            cb_kwargs = dict(name = name, url = url)) 

    def parse_target_pages(self, response, name, url):
        emails = map(lambda x: x.lower(), re.findall(r'[\w\.-]+@[\w\.-]+', response.xpath("/html").extract()[0]))
        emails = set(emails)
        record = CS_FSC_Item()
        record['name'] = name
        record['email'] = emails
        record['url'] = url
        yield record