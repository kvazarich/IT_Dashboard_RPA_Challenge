import logging

from RPA.Browser.Selenium import Selenium

from utils import XlsxSaver, PDFHelper
from agencies_list_parser import AgenciesListParser
from individual_investments_parser import IndividualInvestmentsParser


class ITDashboardRobot:

    def __init__(self):
        self.output_folder = 'output'
        self.list_parser = AgenciesListParser()
        self.detail_parser = IndividualInvestmentsParser('005')
        logfile = f'{self.output_folder}/log_file.log'
        logging.basicConfig(level=logging.INFO, filename=logfile)

    def run(self):
        list_agencies = self.list_parser.parse()
        helper = XlsxSaver(
            values=list_agencies,
            worksheet='Agencies',
            path='output/load_agencies_test.xlsx',
            exclude_keys=['links']
        )
        helper.fill_workbook()
        workbook = helper.get_workbook()
        details = self.detail_parser.parse()
        links = [detail['link'] for detail in details if detail['link']]
        helper = XlsxSaver(
            values=details,
            worksheet='IndividualInvestments',
            path='load_agencies_test.xlsx',
            workbook=workbook,
            exclude_keys=['link']
        )
        helper.fill_workbook()
        helper.save_workbook()
        browser = Selenium()
        filepaths = PDFHelper.load_bulk(links=links, browser=browser, folder_to_load='output')
        PDFHelper.validate(details, filepaths)