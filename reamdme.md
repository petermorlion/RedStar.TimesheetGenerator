# RedStar.TimesheetGenerator

Generates my timesheets based on time entries in (Classic) Freshbooks. 
Should be generic enough so that other timeentry sources and other
timesheet destinations can be added.

## Usage

`dotnet RedStar.TimesheetGenerator.ConsoleApp.dll <username> <auth token> <project id> <YYYYMM> <output.xlsx>`

- username: your Classic Freshbooks account (i.e. the x in x.freshbooks.com)
- auth token: the authentication token you can find under "My Account"
- project id: the project id (go to projects, select the project and copy the number in the URL)
- YYYYMM: the year and month to generate the timesheet for
