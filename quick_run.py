from agencies_list_parser import AgenciesListParser


def print_agencies():
    list_parser = AgenciesListParser()
    agencies = list_parser.parse()
    for agency in agencies:
        print(agency)


if __name__ == "__main__":
    print_agencies()