import requests, zipfile, io


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


x = ZipHelper("http://demografia.stat.gov.pl/bazademografia/Downloader.aspx?file=pl_uro_2012_00_1p.zip&sys=uro")
x.extract_file()
print(x.filename)
