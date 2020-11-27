
import sys


from application import application

global app
app = application()
app.src = "/Users/matt.admin/OneDrive - The Engine Group/clients/Future Borders/HMRC timesheets/HMRC - November 201124.xlsx"
app.master = "/Users/matt.admin/OneDrive - The Engine Group/clients/Future Borders/HMRC timesheets/HMRC billing master.xlsx"

app.clear()
app.convert_excel()
app.get_resources()
app.get_sows()
app.get_engine_timesheet_data()
app.get_hmrc_data()
app.compare_timesheets()
app.write_timesheet_data()
app.terminate()