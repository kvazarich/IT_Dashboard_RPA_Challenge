import logging

from RPA.Excel.Files import Files
from RPA.Browser.Selenium import Selenium
from RPA.FileSystem import FileSystem
import os, re

from RPA.PDF import PDF


class XlsxSaver:
    def __init__(self, values: list, worksheet: str, path: str, workbook=None, exclude_keys=None):
        files = Files()
        if workbook is None:
            self._workbook = files.create_workbook(path=path)
        else:
            self._workbook = workbook
        self._exclude_keys = exclude_keys or []
        self._workbook.create_worksheet(worksheet)
        self._headers_index = {}
        self._worksheet = worksheet
        self._values = values
        self._path = path

    def _fill_headers(self, headers):
        for num, val in enumerate(headers):
            if val not in self._exclude_keys:
                self._headers_index[val] = num + 1
                self._workbook.set_cell_value(1, num + 1, val, self._worksheet)

    def fill_workbook(self):
        self._fill_headers(self._values[0].keys())
        row: dict
        for num, row in enumerate(self._values):
            self._fill_row(row, num + 2)

    def _fill_row(self, row, row_num):
        for key, value in row.items():
            if key not in self._exclude_keys:
                self._workbook.set_cell_value(
                    row_num,
                    self._headers_index[key],
                    value,
                    self._worksheet
                )

    def get_workbook(self):
        return self._workbook

    def save_workbook(self):
        self._workbook.save(self._path)

    def close(self):
        self._workbook.close()
        self._values = None
        self._headers_index = None


class PDFHelper:
    @classmethod
    def load_bulk(cls, links: list, browser: Selenium, folder_to_load=None):
        load_dir = f'{os.getcwd()}/{folder_to_load}' if folder_to_load is not None else os.getcwd()
        browser.set_download_directory(load_dir)
        filepaths = {}
        filesystem = FileSystem()
        for num, link in enumerate(links):
            filename = f"{link.split('/')[-1]}.pdf"
            if filesystem.does_file_exist(f'{load_dir}/{filename}'):
                filesystem.remove_file(f'{load_dir}/{filename}')
                filesystem.wait_until_removed(f'{load_dir}/{filename}')
            browser.open_available_browser(link)
            browser.wait_until_element_is_visible(locator='css:div#business-case-pdf a')
            browser.click_element(locator='css:div#business-case-pdf a')
            filepaths[link] = f'{load_dir}/{filename}'
            filesystem.wait_until_created(
                f'{load_dir}/{filename}',
                timeout=60.0*5
            )
            browser.close_all_browsers()
        return filepaths

    @classmethod
    def parse(cls, filepath):
        pdf = PDF()
        text = pdf.get_text_from_pdf(filepath)
        for _, string_ in text.items():
            if 'Section A' in string_:
                if 'Section B' in string_:
                    section_A = string_.split('Section B')[0]
                else:
                    section_A = string_
                match_obj = re.search(
                    'Name of this Investment\: ([\n\d()A-Za-z \,\-]+).*2. Unique Investment Identifier .UII.: ([\d\- ]+)',
                    section_A,
                    re.DOTALL
                )
                if match_obj is None or match_obj.group(2) is None:
                    logging.error(msg=f'name or uii not found in {filepath}')
                name = match_obj.group(1).strip()
                uii = match_obj.group(2).strip()
                return name, uii

    @classmethod
    def validate(cls, individual_investments, filepaths):
        for individual_investment in individual_investments:
            if individual_investment['link']:
                filepath = filepaths[individual_investment['link']]
                name_of_this_investment, UII = cls.parse(filepath)
                if name_of_this_investment != individual_investment['Investment Title']:
                    logger = logging.getLogger()
                    logger.info(
                        msg=f'{name_of_this_investment}) not equal {individual_investment["Investment Title"]} link: {individual_investment["link"]}'
                    )


