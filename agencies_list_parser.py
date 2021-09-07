from RPA.Browser.Selenium import Selenium


class AgenciesListParser:
    URL = "https://itdashboard.gov/drupal/#home-dive-in"

    def __init__(self):
        self.browser = Selenium()
        self.browser.open_available_browser(self.URL)

    def _get_agencies(self):
        try:
            self.browser.wait_until_element_is_visible(locator='css:div#agency-tiles-2-container>div>div>div>div')
            agencies = self.browser.get_webelements(locator='css:div#agency-tiles-2-container>div>div>div>div')
        except:
            raise
        return agencies

    def _get_name(self, agency_element):
        self.browser.wait_until_element_is_visible(locator=[agency_element, 'css:span:nth-of-type(1)'])
        return self.browser.get_webelement(locator=[agency_element, 'css:span:nth-of-type(1)']).text

    def _get_amount(self, agency_element):
        self.browser.wait_until_element_is_visible(locator=[agency_element, 'css:span:nth-of-type(2)'])
        return self.browser.get_webelement(locator=[agency_element, 'css:span:nth-of-type(2)']).text

    def _get_link(self, agency_element):
        self.browser.wait_until_element_is_visible(locator=[agency_element, 'css:a'])
        return self.browser.get_element_attribute(locator=[agency_element, 'css:a'], attribute='href')

    def parse(self):
        agencies_elements = self._get_agencies()
        agencies = []
        for agency_element in agencies_elements:
            agency = {
                'name': self._get_name(agency_element),
                'amount': self._get_amount(agency_element),
                'link': self._get_link(agency_element)
            }
            agencies.append(agency)
        return agencies
