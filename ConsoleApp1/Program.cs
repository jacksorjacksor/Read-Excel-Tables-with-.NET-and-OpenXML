using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

// Combination of these articles:
// > How to get all rows/cell values from a range with OpenXml and c#
// https://stackoverflow.com/questions/54483407/how-to-get-all-rows-cell-values-from-a-range-with-openxml-and-c-sharp

// > Retrieve the values of cells in a spreadsheet document (Open XML SDK)
// https://docs.microsoft.com/en-us/office/open-xml/how-to-retrieve-the-values-of-cells-in-a-spreadsheet

const string filePath = "E:\\ExcelTestFile.xlsx";

using (var document = SpreadsheetDocument.Open(filePath, true))
{
    string tableRange = "";

    WorkbookPart wbPart = document.WorkbookPart;
    WorksheetPart wsPart = wbPart.WorksheetParts.First();
    Worksheet worksheet = wbPart.WorksheetParts.First().Worksheet;
    Sheet sheet = GetSheetFromWorkSheet(wbPart, wsPart);

    // Takes FIRST sheetName in document - if more than will need to loop:
    // // https://stackoverflow.com/questions/7504285/how-to-retrieve-tab-names-from-excel-sheet-using-openxml

    if (wbPart != null)
        foreach (WorksheetPart ws in wbPart.WorksheetParts)
        {
            foreach (var table in ws.TableDefinitionParts)
            {
                tableRange = table.Table.Reference;
            }
        }

    // Get table range = able.Table.Reference
    // Regex splits Excel range into groups:
    string Pattern = @"([A-Z]*)([0-9]*):([A-Z]*)([0-9]*)";
    Match regexForTableRange = Regex.Match(tableRange, Pattern);

    // Columns in Excel are alphabetical - converting to int
    var colStart = StringToInt(regexForTableRange.Groups[1].Value);
    var rowStart = int.Parse(regexForTableRange.Groups[2].Value);
    var colEnd = StringToInt(regexForTableRange.Groups[3].Value);
    var rowEnd = int.Parse(regexForTableRange.Groups[4].Value);

    // Aim:
    // -- for each ROW, go over each COLUMN

    // ROW
    for (var rowIndex = rowStart+1; rowIndex < rowEnd + 1; rowIndex++)
    {
        // COLUMN
        for (var colIndex = colStart; colIndex < colEnd + 1; colIndex++)
        {
            var cellRef = IntToString(colIndex-1) + rowIndex;
            var value = GetCellValue(document, sheet.Name, cellRef);
            Console.WriteLine(value);
        }
    }
}


// UTIL FUNCTIONS

static string IntToString(int index)
{
    // https://stackoverflow.com/questions/10373561/convert-a-number-to-a-letter-in-c-sharp-for-use-in-microsoft-excel
    const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    var value = "";

    if (index >= letters.Length)
        value += letters[index / letters.Length - 1];

    value += letters[index % letters.Length];

    return value;
}

static int StringToInt(string columnName)
{
    // https://stackoverflow.com/questions/667802/what-is-the-algorithm-to-convert-an-excel-column-letter-into-its-number
    if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

    columnName = columnName.ToUpperInvariant();

    int sum = 0;

    for (int i = 0; i < columnName.Length; i++)
    {
        sum *= 26;
        sum += (columnName[i] - 'A' + 1);
    }
    return sum;
}


static Sheet GetSheetFromWorkSheet(WorkbookPart workbookPart, WorksheetPart worksheetPart)
{
    // https://stackoverflow.com/questions/7504285/how-to-retrieve-tab-names-from-excel-sheet-using-openxml
    string relationshipId = workbookPart.GetIdOfPart(worksheetPart);
    IEnumerable<Sheet> sheets = workbookPart.Workbook.Sheets.Elements<Sheet>();
    return sheets.FirstOrDefault(s => s.Id.HasValue && s.Id.Value == relationshipId);
}


static string GetCellValue(SpreadsheetDocument document, string sheetName, string addressName)
{
    // https://docs.microsoft.com/en-us/office/open-xml/how-to-retrieve-the-values-of-cells-in-a-spreadsheet
    string value = null;

    // Open the spreadsheet document for read-only access.
    // Retrieve a reference to the workbook part.
    WorkbookPart wbPart = document.WorkbookPart;

    // Find the sheet with the supplied name, and then use that
    // Sheet object to retrieve a reference to the first worksheet.
    Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

    // Throw an exception if there is no sheet.
    if (theSheet == null)
    {
        throw new ArgumentException("sheetName");
    }

    // Retrieve a reference to the worksheet part.
    WorksheetPart wsPart =
        (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

    // Use its Worksheet property to get a reference to the cell
    // whose address matches the address you supplied.
    Cell theCell = wsPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == addressName);

    // If the cell does not exist, return an empty string.
    if (theCell.InnerText.Length > 0)
    {
        value = theCell.InnerText;

        // If the cell represents an integer number, you are done.
        // For dates, this code returns the serialized value that
        // represents the date. The code handles strings and
        // Booleans individually. For shared strings, the code
        // looks up the corresponding value in the shared string
        // table. For Booleans, the code converts the value into
        // the words TRUE or FALSE.
        if (theCell.DataType != null)
        {
            switch (theCell.DataType.Value)
            {
                case CellValues.SharedString:

                    // For shared strings, look up the value in the
                    // shared strings table.
                    var stringTable =
                        wbPart.GetPartsOfType<SharedStringTablePart>()
                            .FirstOrDefault();

                    // If the shared string table is missing, something
                    // is wrong. Return the index that is in
                    // the cell. Otherwise, look up the correct text in
                    // the table.
                    if (stringTable != null)
                    {
                        value =
                            stringTable.SharedStringTable
                                .ElementAt(int.Parse(value)).InnerText;
                    }

                    break;

                case CellValues.Boolean:
                    switch (value)
                    {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }

                    break;
            }
        }
    }
    return value;
}