using System.Data;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelETLHelper;

/// <summary>
/// <para>Excel helper language extension, perform commond operation for Excel files with optimized performances.</para>
/// <para>Example:</para>
/// <para>".ExportToXls()" can write a DataTable or DataSet as Xls to a destination path.</para>
/// </summary>
public static class ExtRefOpenXml
{
    #region DataSet
    /// <summary>
    /// Export the DataTable to an .xls file
    /// </summary>
    public static bool ExportToXls(this DataTable dt, string destinationPath) => new DataSet() { Tables = { dt } }.ExportToXls(destinationPath);

    /// <summary>
    /// Export the DataSet to an .xls file
    /// </summary>
    public static bool ExportToXls(this DataSet ds, string destinationPath)
    {
        using (var workbook = SpreadsheetDocument.Create(destinationPath, SpreadsheetDocumentType.Workbook))
        {
            workbook.AddWorkbookPart();
            if (workbook == null || workbook.WorkbookPart == null)
                return false;
            workbook.WorkbookPart.Workbook = new Workbook { Sheets = new Sheets() };

            foreach (DataTable table in ds.Tables)
            {
                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                sheetPart.Worksheet = new Worksheet(sheetData);

                var sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                if (sheets == null)
                    return false;
                var relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);
                var sheetId = !sheets.Elements<Sheet>().Any() ? 1 : sheets.Elements<Sheet>().Max(s => s?.SheetId?.Value) + 1;
                var sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = table.TableName };
                sheets.Append(sheet);

                var headerRow = new Row();
                var columns = new List<string>();
                foreach (DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    var cell = new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(column.ColumnName)
                    };

                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {
                    var newRow = new Row();
                    foreach (var col in columns)
                    {
                        if (!dsrow.Table.Columns.Contains(col) || dsrow[col] == null || dsrow[col].GetType() == typeof(DBNull))
                            continue;
                        var cell = new Cell
                        {
                            DataType = CellValues.String,
#pragma warning disable CS8604 // pragma "possible null reference": addressed in the "if (...) continue;" above.
                            CellValue = new CellValue(dsrow[col].ToString())
#pragma warning restore CS8604
                        };

                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }
            }
        }

        return true;
    }
    #endregion DataSet
}