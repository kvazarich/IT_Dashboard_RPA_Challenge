from RPA.Browser.Selenium import Selenium

from agencies_list_parser import AgenciesListParser
from individual_investments_parser import IndividualInvestmentsParser
from utils import XlsxSaver, PDFHelper


def get_agencies():
    list_parser = AgenciesListParser()
    return list_parser.parse()


def get_agencies_workbook(agencies):
    helper = XlsxSaver(values=agencies, worksheet='Agencies', path='load_agencies_test.xlsx')
    helper.fill_workbook()
    return helper.get_workbook()


def save_workbook(data, workbook):
    helper = XlsxSaver(
        values=data,
        worksheet='IndividualInvestments',
        path='load_agencies_test.xlsx',
        workbook=workbook
    )
    helper.fill_workbook()
    helper.save_workbook()


def get_agency_details(agency_id):
    detail_parser = IndividualInvestmentsParser(agency_id)
    return detail_parser.parse()


if __name__ == "__main__":
    # agencies = get_agencies()
    # wb = get_agencies_workbook(agencies)
    details = get_agency_details(agency_id="005")
    links = [detail['link'] for detail in details if detail['link']]
    # save_workbook(details, wb)
    browser = Selenium()
    filepaths = PDFHelper.load_bulk(links=links, browser=browser, folder_to_load='output')
    # print(PDFHelper.parse('/home/kvazarich/PycharmProjects/IT_Dashboard_RPA_Challenge/output/005-000000018.pdf'))
    PDFHelper.validate(details, filepaths)
