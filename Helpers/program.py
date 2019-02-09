import requests, zipfile, io, xlrd


class ZipHelper:

    def __init__(self, Http):
        request = requests.get(Http)
        file = zipfile.ZipFile(io.BytesIO(request.content), 'r')
        self.file = file
        self.filename = ''
        self.pathToExtract = '../Files'

    def extract_file(self):
        self.file.extractall(self.pathToExtract)
        for info in self.file.infolist():
            self.filename = info.filename
        return self.filename


class ExcelReader:
    def __init__(self):
        self.pages = {
            2012: "http://demografia.stat.gov.pl/bazademografia/Downloader.aspx?file=pl_uro_2012_00_1p.zip&sys=uro",
            2013: "http://demografia.stat.gov.pl/bazademografia/Downloader.aspx?file=pl_uro_2013_00_1p.zip&sys=uro",
            2014: "http://demografia.stat.gov.pl/bazademografia/Downloader.aspx?file=pl_uro_2014_00_1p.zip&sys=uro",
            2015: "http://demografia.stat.gov.pl/bazademografia/Downloader.aspx?file=pl_uro_2015_00_1p.zip&sys=uro",
            2016: "http://demografia.stat.gov.pl/bazademografia/Downloader.aspx?file=pl_uro_2016_00_1p.zip&sys=uro",
            2017: "http://demografia.stat.gov.pl/bazademografia/Downloader.aspx?file=pl_uro_2017_00_1p.zip&sys=uro",
        }
        self.sheet_name = "Ogółem"
        self.necessary_data = dict()

    def get_cell_coordinates(self, year):
        if year < 2015:
            return {'x': 21, 'y': 5}
        return {'x': 18, 'y': 5}

    def create_necessary_data(self):
        print('Downloading files, please be patient.')
        for year in self.pages:
            zipHelper = ZipHelper(self.pages[year])
            zipHelper.extract_file()
            workbook = xlrd.open_workbook(zipHelper.pathToExtract + "/" + zipHelper.filename, on_demand=True)
            sheet = workbook.sheet_by_name(self.sheet_name)
            coordinates = self.get_cell_coordinates(year)
            self.necessary_data[year] = sheet.cell(coordinates['x'], coordinates['y']).value


class Interpreter:

    def __init__(self, data):
        key_min = min(data.keys(), key=(lambda k: data[k]))
        self.date_with_min_value = data[key_min]
        self.data = data

    def compare_two_years_before_and_after_500_plus(self):
        average_before_500_plus = (self.data[2014] + self.data[2015]) / 2
        average_after_500_plus = (self.data[2016] + self.data[2017]) / 2
        if average_before_500_plus > average_after_500_plus:
            print("Average of births 2 years before is higher then 2 years after of inaugurate 500 plus.")
        if average_before_500_plus < average_after_500_plus:
            print("Average of births 2 years after is higher then 2 years after of inaugurate 500 plus.")
        if average_before_500_plus == average_after_500_plus:
            print("Average of births 2 years before before is a same 2 years after inaugurate 500 plus.")

    def information_about_year_when_500_plus_inaugurated(self):
        if self.data[2015] > self.data[2016]:
            print('In year when 500 plus program has been inaugurated was less number of births then in year before it.')
        if self.data[2015] < self.data[2016]:
            print('In year when 500 plus program has been inaugurated was more number of births then in year before it.')
        if self.data[2015] == self.data[2016]:
            print('In year when 500 plus program has been inaugurated was same number of births then in year before it.')
        if self.data[2016] > self.data[2017]:
            print('In year when 500 plus program has been inaugurated was more number of births then in year after it.')
        if self.data[2016] < self.data[2017]:
            print('In year when 500 plus program has been inaugurated was less number of births then in year after it.')
        if self.data[2016] == self.data[2017]:
            print('In year when 500 plus program has been inaugurated was same number of births then in year after it.')


reader = ExcelReader()
reader.create_necessary_data()
print(reader.necessary_data)

interpreter = Interpreter(reader.necessary_data)
interpreter.compare_two_years_before_and_after_500_plus()
interpreter.information_about_year_when_500_plus_inaugurated()
