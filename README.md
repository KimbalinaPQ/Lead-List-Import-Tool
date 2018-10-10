# Lead-List-Import-Tool
The **Lead List Import Tool** (aka **IMC List Import Tool** or **Marketing List Import/Upload Tool**) is used by the Marketing team to push a list of contacts ("leads") typically gathered an event such as a trade show to the Marketing Cloud leads database.

The app is an ASPX file hosted in [myProQuest](https://myproquest.sharepoint.com/) at [https://myproquest.sharepoint.com/teams/Marketing/ops/IMC%20List%20Import%20Tool/Forms/AllItems.aspx](https://myproquest.sharepoint.com/teams/Marketing/ops/IMC%20List%20Import%20Tool/Forms/AllItems.aspx).  To deploy, push updated project files to this directory.

Users access the app from a link in the [Marketing Operations homepage](https://myproquest.sharepoint.com/teams/Marketing/ops/SitePages/Home.aspx).

Direct link to the app is [https://myproquest.sharepoint.com/teams/Marketing/ops/IMC%20List%20Import%20Tool/TSUpload-Start.aspx](https://myproquest.sharepoint.com/teams/Marketing/ops/IMC%20List%20Import%20Tool/TSUpload-Start.aspx).

## Usage
The marketer enters the leads information into an Excel spreadsheet with the name `lead-template.xlsx`.  The empty template spreadsheet is available to users at the [hosting page](https://myproquest.sharepoint.com/teams/Marketing/ops/IMC%20List%20Import%20Tool/Forms/AllItems.aspx).  The leads sheet must use this format.

The user runs the app and selects an "Upload Type" which determines the destination for the data.  Then they select the file for import.  The app scans the file for data erros and lists them on the page.  If the data does not meet all criteria for correctness, the file is rejected.  The user must fix the error(s) in Excel and retry the import.



## Operation

### Upload Types

### Criteria

### Sending email alerts and creating tasks

## Dependencies

----

This program was created by Joe Insko using the `js-xslx` library ([SheetJS.com](http://sheetjs.com)).  First updated by Jim Saiya (@jsaiyapq) in October 2017.
