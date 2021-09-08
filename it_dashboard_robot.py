from utils import XlsxSaver
from agencies_list_parser import AgenciesListParser
from individual_investments_parser import IndividualInvestmentsParser


class ITDashboardRobot:
    SAVER_CLASS = XlsxSaver

    def __init__(self):
        self.list_parser = AgenciesListParser()
        self.detail_parser = IndividualInvestmentsParser('005')

    def run(self):
        pass