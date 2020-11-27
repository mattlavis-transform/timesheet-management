U# Timesheet extracter

## Implementation steps

- Create and activate a virtual environment, e.g.

  `python3 -m venv venv/`
  `source venv/bin/activate`

- Install necessary Python modules 

  - et-xmlfile==1.0.1
  - jdcal==1.4.1
  - numpy==1.19.4
  - openpyxl==3.0.5
  - pandas==1.1.4
  - python-dateutil==2.8.1
  - pytz==2020.4
  - six==1.15.0
  - xlrd==1.2.0

  via `pip3 install -r requirements.txt`


## Usage

To extract only the relevant information from the full timesheet document:

`python3 main.py stw november full`
`python3 main.py smartfreight november full`
`python3 main.py ott november full`
