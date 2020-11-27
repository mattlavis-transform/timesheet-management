import sys
from datetime import datetime


class hmrc_timesheet_entry(object):
    def __init__(self):
        self.project_id = ""
        self.resource_name = ""
        self.resource_type = ""
        self.day_date = ""
        self.day_date_as_date = None
        self.timesheet_period = ""
        self.timesheet_period_formatted = ""
        self.day_of_week = ""
        self.month = ""
        self.day_date_formatted = ""
        self.day_rate = None
        self.hours = None
        self.sow_id = ""
        self.resource = None

    def get_resource_name(self, app, col):
        for res in app.resources:
            if res.match_column == col:
                self.resource_name = res.resource_name
                self.resource_type = res.resource_type
                self.resource = res
                break

    def cost(self):
        if self.resource_type == "Employee":
            s = self.hours_fixed * self.day_rate / 7.5
        else:
            s = self.hours_fixed * self.day_rate / 8
        return s

    def process_date(self):
        self.day_date_as_date = datetime.strptime(self.day_date, '%Y-%m-%d')
        self.day_date_formatted = self.day_date_as_date.strftime("%d/%m/%Y")
        self.day_date_formatted = datetime.strptime(
            self.day_date_formatted, "%d/%m/%Y")

        self.day_of_week = self.day_date_as_date.strftime("%A")
        self.month = self.day_date_as_date.strftime("%B")

        date_time_obj = datetime.strptime(self.timesheet_period, '%Y-%m-%d')
        self.timesheet_period_formatted = date_time_obj.strftime("%d/%m/%Y")
        self.timesheet_period_formatted = datetime.strptime(
            self.timesheet_period_formatted, "%d/%m/%Y")

    def fix_hours(self):
        if self.hours == "":
            self.hours = 0
        self.hours = float(self.hours)

    def get_day_rate(self, resources):
        for res in resources:
            if res.resource_name == self.resource_name:
                self.day_rate = res.day_rate

    def get_sow(self, sows):
        for my_sow in sows:
            if my_sow.project_id == self.project_id:
                if my_sow.start_date <= self.day_date_as_date and my_sow.end_date >= self.day_date_as_date:
                    self.sow_id = my_sow.sow_id
                    break

    def __str__(self):
        s = self.project_id + ", "
        s += self.resource_name + ", "
        s += self.resource_type + ", "
        s += self.day_date + ", "
        s += str(self.hours)

        return (s)
