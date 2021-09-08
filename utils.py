from RPA.Excel.Files import Files
from RPA.Browser.Selenium import Selenium
from RPA.FileSystem import FileSystem
import os

class XlsxSaver:
    def __init__(self, values: list, worksheet: str, path: str, workbook=None):
        files = Files()
        if workbook is None:
            self._workbook = files.create_workbook(path=path)
        else:
            self._workbook = workbook
        self._workbook.create_worksheet(worksheet)
        self._headers_index = {}
        self._worksheet = worksheet
        self._values = values
        self._path = path

    def _fill_headers(self, headers):
        for num, val in enumerate(headers):
            self._headers_index[val] = num + 1
            self._workbook.set_cell_value(1, num + 1, val, self._worksheet)

    def fill_workbook(self):
        self._fill_headers(self._values[0].keys())
        row: dict
        for num, row in enumerate(self._values):
            self._fill_row(row, num + 2)

    def _fill_row(self, row, row_num):
        for key, value in row.items():
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


class PDFLoader:
    @classmethod
    def load_bulk(cls, links: list, browser: Selenium, folder_to_load=None):
        load_dir = f'{os.getcwd()}/{folder_to_load}' if folder_to_load is not None else os.getcwd()
        browser.set_download_directory(load_dir)
        filenames = []
        filesystem = FileSystem()
        for num, link in enumerate(links):
            filename = f"{link.split('/')[-1][:-1]}.pdf"
            if filesystem.does_file_exist(f'{load_dir}/{filename}'):
                filesystem.remove_file(f'{load_dir}/{filename}')
                filesystem.wait_until_removed(f'{load_dir}/{filename}')
            browser.open_available_browser(link)
            browser.wait_until_element_is_visible(locator='css:div#business-case-pdf a')
            browser.click_element(locator='css:div#business-case-pdf a')
            filenames.append(filename)
            filesystem.wait_until_created(
                f'{load_dir}/{filename}',
                timeout=60.0*5
            )
            browser.close_all_browsers()
        return filenames

    @classmethod
    def parse(cls, filename):
        pass


class ElementNotLoaded(Exception):
    pass