// See https://aka.ms/new-console-template for more information

using ClosedXML.Excel;
using ClosedXML.Report;
using ClosedXmlTest.Models;

const string TEMPLATE_FILE = @".\template.xlsx";
const string OUTPUT_FILE = @".\output.xlsx";
const string TEMPLATE_WORKSHEET_NAME = "Sheet1";
const string DATA_RANGE_NAME = "DataItems"; // Range already defined in template
const string HEADERS_RANGE_NAME = "ReportHeaders"; // Range already defined in template
const string REPORT_TABLE_RANGE_NAME = "ReportTable"; // Range will be created programatically

var reportContent = new 
{
    DataItems = new List<DataItem>()
};

// Create some data to fill the report template with
var rowsToGenerate = 20;
var random = new Random(DateTime.Now.Millisecond);

for (int dataItemCount = 0; dataItemCount < rowsToGenerate; dataItemCount++)
{
    var wordCount = random.Next(1, 20);
    var wrappedText = string.Join(" ", Enumerable.Repeat("text", wordCount));

    var lineCount = random.Next(1, 3);
    var multilineText = string.Join(";\r\n", Enumerable.Repeat("text", lineCount));

    var dataItem = new DataItem 
    { 
        Id = dataItemCount,
        WrappedText = wrappedText,
        MultilineText = multilineText 
    };

    reportContent.DataItems.Add(dataItem);
}

// Apply the data to the template
var template = new XLTemplate(TEMPLATE_FILE);
template.AddVariable(reportContent);
template.Generate();

// Remove any explicit heights of rows that may have wrapped text so they automatically
// adjust to the required height for their content.
var rows = template.Workbook.Range(DATA_RANGE_NAME).Rows();

foreach (var row in rows)
{
    row.WorksheetRow().ClearHeight();
}

// Format the data as a table, because rows of variable height are hard to read. Unfortunately
// this can't be done in the template itself, as new rows are not automatically added to an
// existing table range. 
var headerTopLeftCell = template.Workbook.Range(HEADERS_RANGE_NAME).Cells().First();
var dataBottomRightCell = template.Workbook.Range(DATA_RANGE_NAME).Cells().Last();

template
    .Workbook
    .Worksheet(TEMPLATE_WORKSHEET_NAME)
    .Range(headerTopLeftCell.Address, dataBottomRightCell.Address)
    .AddToNamed(REPORT_TABLE_RANGE_NAME);

template.Workbook.Range(REPORT_TABLE_RANGE_NAME).CreateTable();

// Delete the service column - the leftmost column in templates is used for tags that affect
// report generation and is automatically cleared, but the empty column is not automatically
// deleted. No configuration tags are used in the example template.
template.Workbook.Worksheet(TEMPLATE_WORKSHEET_NAME).Columns().First().Delete();

template.SaveAs(OUTPUT_FILE);