# RedStar.TimesheetGenerator

Generates my timesheets based on time entries in Freshbooks. 
Should be generic enough so that other timeentry sources and other
timesheet destinations can be added.

## Usage

```
SET freshbooks_api_client_id=...
SET freshbooks_api_client_secret=...
SET freshbooks_client_id=...
SET freshbooks_business_id=...
".\RedStar.TimesheetGenerator.ConsoleApp.exe" freshbooks miaa 202106 ".\2021-06.xlsx"
```

The above generates a timesheet for June 2021 using a the `freshbooks` input plugin and the `miaa` output plugin and writes it to the given Excel file.

The environment variables are `freshbooks` plugin specific:

 - `freshbooks_api_client_id`: the client ID to connect to the Freshbooks API
 - `freshbooks_api_client_secret`: the client secret to connect to the Freshbooks API
 - `freshbooks_client_id`: the ID of the client in my Freshbooks account
 - `freshbooks_business_id`: the ID of the business in my Freshbooks account
