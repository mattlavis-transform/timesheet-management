import sys
from datetime import datetime


class sow(object):
    def __init__(self):
        self.project_id = ""
        self.sow_id = ""
        self.start_date = None
        self.end_date = None

    def format_dates(self):
        self.start_date = datetime.strptime(self.start_date, '%d/%m/%Y')
        self.end_date = datetime.strptime(self.end_date, '%d/%m/%Y')
        
    def __str__(self):
        s = self.project_id + ", "
        s += self.sow_id + ", "
        s += str(self.start_date) + ", "
        s += str(self.end_date)
        return (s)
