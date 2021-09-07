import datetime

from RPA.Browser.Selenium import Selenium


class IndividualInvestmentsParser:
    BASE_URL = "https://itdashboard.gov/drupal/summary/{}"

    def __init__(self, agency_id):
        self._link = self.BASE_URL.format(agency_id)
        self.browser = Selenium()
        self.browser.open_available_browser(self._link)

    def _get_investments_table(self):
        self.browser.wait_until_element_is_visible(
            locator='css:div#investments-table-widget div.pageSelect select option:nth-of-type(4)',
            timeout=datetime.timedelta(minutes=1)
        )
        self.browser.click_element(
            locator='css:div#investments-table-widget div.pageSelect select option:last-of-type'
        )
        self.browser.wait_until_element_is_visible(
            locator='css:div#investments-table-widget table#investments-table-object tbody tr:nth-of-type(11) td',
            timeout=datetime.timedelta(minutes=1)
        )
        agencies = self.browser.get_webelement(
            locator='css:div#investments-table-widget'
        )
        return agencies

    def _get_table_headers(self, table):
        self.browser.wait_until_element_is_visible(
            locator=[table, 'css:div.dataTables_scrollHead table th'],
            timeout=datetime.timedelta(minutes=1)
        )
        elements = self.browser.get_webelements(locator=[table, 'css:div.dataTables_scrollHead table th'])
        return [element.text for element in elements]

    def _get_table_rows(self, table):
        self.browser.wait_until_element_is_enabled(locator=[table, 'css:table#investments-table-object tbody tr'])
        elements = self.browser.get_webelements(locator=[table, 'css:table#investments-table-object tbody tr'])
        return elements

    def _get_row_cells(self, table_row):
        elements = self.browser.get_webelements(locator=[table_row, 'css:td'])
        return elements

    def parse(self):
        investments_table = self._get_investments_table()
        headers = self._get_table_headers(table=investments_table)
        rows = self._get_table_rows(table=investments_table)
        investments = []
        for row in rows:
            investment = {}
            for num, cell in enumerate(self._get_row_cells(row)):
                count_a = self.browser.get_element_count(locator=[cell, 'css:a'])
                if count_a > 0:
                    investment['link'] = self.browser.get_element_attribute(locator=[cell, 'css:a'], attribute='href')
                else:
                    investment['link'] = ''
                investment[headers[num]] = cell.text
            investments.append(investment)
        return investments
