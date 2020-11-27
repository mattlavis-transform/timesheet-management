import sys


class resource(object):
    def __init__(self):
        self.project_id = ""
        self.resource_name = ""
        self.resource_type = ""
        self.day_rate = None
        self.match_column = None
        self.charge_under = ""
        self.hours_per_day = None

    def get_hours_per_day(self):
        if self.resource_type == "Employee":
            self.hours_per_day = 7.5
        else:
            self.hours_per_day = 8
        
    def __str__(self):
        s = self.project_id + ", "
        s += self.resource_name + ", "
        s += self.resource_type + ", "
        s += str(self.day_rate) + ", "
        s += self.charge_under + ", "
        s += str(self.hours_per_day) + ", "
        s += str(self.match_column)
        return (s)
