
import sys


from application import application

global app
app = application()
app.get_config()

app.clear()
app.convert_excel()
app.get_resources()
app.get_sows()
app.get_engine_timesheet_data()
# app.get_hmrc_data()
# app.compare_timesheets()
app.write_timesheet_data()
app.terminate()