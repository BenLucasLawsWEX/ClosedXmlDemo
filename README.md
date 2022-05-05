# ClosedXmlDemo
Demo project for ClosedXml and ClosedXmlReport libraries

## Usage
Build an run the application. A report 'output.xlsx' will be generated in the same folder as the executable.

## What does the application do?
* Generates 20 rows of data with text content of a random length
* Replaces placeholder tokens in the report template 'template.xlsx'
* Automatically resizes rows to vertically fit word wrapped content
* Formats the report data as a table
* Saves the report as a new file 'output.xlsx'

## Notes
Report generation will fail if the output report is already open in Excel.
