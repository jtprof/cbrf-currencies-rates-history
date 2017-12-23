from base_client import BaseClient
from datetime import date
from datetime import datetime
import xml.etree.ElementTree as ET
from collections import OrderedDict
import string

class CBRFRateByDateClient(BaseClient):
    _base_URL = 'http://www.cbr.ru/scripts/XML_daily.asp?date_req='
    _cbrf_date_format = '%d/%m/%Y'
    _iso_date_format = '%Y%m%d'
    date = datetime.now()
    cbrfdate = date.strftime(_cbrf_date_format)
    isodate =  date.strftime(_iso_date_format)

    method = 'users.get'
    rates = []

    def __init__(self, date):
        super(BaseClient, self).__init__()
        self.setdate(date)
        self.rates = []

    def setdate(self, date):
        self.cbrfdate = date.strftime(self._cbrf_date_format)
        self.isodate = date.strftime(self._iso_date_format)
        self.date = date

    def generate_url(self, method):
        url = '{0}{1}'.format(self._base_URL, self.cbrfdate)
        return url

    def response_handler(self, response):
        try:
            root = ET.fromstring(response.text.encode('latin1'))
            for c1 in root:
                cur = OrderedDict()
                for c2 in c1:
                    cur.update({c2.tag :  c2.text})
                if cur.has_key('Value'):
                    v = float(string.replace(cur['Value'], ',', '.'))
                    if cur.has_key('Nominal'):
                         v = v / float(string.replace(cur['Nominal'], ',', '.'))
                    cur['Value'] = v
                self.rates.append(cur)
        finally:
            pass
        return response
