using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace XlsToJson
{
    internal static class XlsToJsonService
    {
        internal static JObject ProcessDocument(WorkbookPart workbookPart, Regex[]? filters, bool excludeHiddenRows, bool excludeHiddenColumns, out string cellsExceptions)
        {
            cellsExceptions = string.Empty;

            if (workbookPart?.Workbook?.DefinedNames == null) return new JObject();

            var filteredDefinedNames = FilterDefinedNamesWithRegex(GetDefinedNames(workbookPart), filters);
           
            if (!filteredDefinedNames.Any()) return new JObject();

            var filteredSheets = GetSheetsForFilteredDefinedNames(workbookPart, filteredDefinedNames);

            if (filteredSheets == null || !filteredSheets.Any()) return new JObject();

            var definedNameValuePairs = ExtractDefinedNameValuePairs(workbookPart, filteredDefinedNames, filteredSheets, excludeHiddenRows, excludeHiddenColumns, out StringBuilder cellsExceptionMessages);

            cellsExceptions = cellsExceptionMessages.ToString();

            return definedNameValuePairs.Any()
                ? JObject.Parse(JsonConvert.SerializeObject(definedNameValuePairs))
                : new JObject();
        }

        internal static List<DefinedName> GetDefinedNames(WorkbookPart workbookPart)
        {
            var definedNames = workbookPart.Workbook.Descendants<DefinedName>().Where(df => !string.Equals(df.InnerText, "#REF!", StringComparison.OrdinalIgnoreCase)).ToList();

            return definedNames;
        }
        private static Dictionary<string, JToken> ExtractDefinedNameValuePairs(WorkbookPart workbookPart, Dictionary<string, string> filteredDefinedNames,
                                                                               Dictionary<string, string> filteredSheets, bool excludeHiddenRows, bool excludeHiddenColumns, out StringBuilder cellsExceptionMessages)
        {
            var results = new Dictionary<string, JToken>();
            cellsExceptionMessages = new StringBuilder();

            foreach (var (key, value) in filteredSheets)
            {
                var sheetData = GetSheetDataByRelationshipId(workbookPart, value);

                if (sheetData == null) continue;

                var matchingNames = filteredDefinedNames
                    .Where(fdn => fdn.Value.Contains(key))
                    .Where(fdn => fdn.Value.Contains("$"))
                    .ToList();

                foreach (var (name, cellReference) in matchingNames)
                {
                    try
                    {
                        var columnName = cellReference.Split('$', '$')[1];

                        var rowIndex = uint.Parse(cellReference[(cellReference.LastIndexOf("$", StringComparison.Ordinal) + 1)..]);

                        var cell = GetCell(sheetData, columnName, rowIndex, excludeHiddenRows, excludeHiddenColumns);

                        if (cell == null) continue;

                        var cellValue = GetCellValue(workbookPart, cell);

                        if (cellValue != null)
                        {
                            results[name] = cellValue;
                        }
                    }
                    catch(Exception ex)
                    {
                        cellsExceptionMessages.AppendLine($"Error occurred in cell {name} from location {cellReference} : {ex.Message} ");
                    }
                }
            }
            return results;
        }

        internal static Dictionary<string, string> FilterDefinedNamesWithRegex(List<DefinedName> definedNames, Regex[]? filters)
        {
            var filteredDefinedNames = new Dictionary<string, string>();

            foreach (var definedName in definedNames)
            { 
                if (definedName.Name == null || definedName.Name.Value == null)
                {
                    continue;
                }
                if (filters == null && !filteredDefinedNames.ContainsKey(definedName.Name.Value))
                {
                    filteredDefinedNames.Add(definedName.Name.Value, definedName.InnerText);
                }
                else if (filters != null)
                {
                    foreach (var filter in filters)
                    {
                        if (filter.IsMatch(definedName.Name.Value) && !filteredDefinedNames.ContainsKey(definedName.Name.Value))
                        {
                            filteredDefinedNames.Add(definedName.Name.Value, definedName.InnerText);
                        }
                    }
                }
            }
            return filteredDefinedNames;
        }

        internal static JToken? GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            JToken? cellValue = null;
           
                if (cell.DataType != null)
                {
                    if (cell.DataType != CellValues.SharedString) return cellValue;

                    if (!int.TryParse(cell.InnerText, out var id)) return cellValue;

                    var item = GetSharedStringItemById(workbookPart, id);

                    if (item == null)
                    {
                        return cellValue;
                    }
                    cellValue = item.Text != null ? item.Text.Text : item.InnerText;
                }
                else if (cell.StyleIndex! != null && cell.CellValue != null && CheckIfFormatIsDate(workbookPart, cell))
                {
                    cellValue = cell.CellValue.Text != "0" ? DateTime.FromOADate(Convert.ToDouble(cell.CellValue.Text)).ToShortDateString() : cell.CellValue.Text;
                }
                else if (cell.StyleIndex! != null && cell.CellValue != null && CheckIfFormatIsNumber(workbookPart, cell))
                {
                    decimal.TryParse(cell.CellValue.Text, out var result);

                    if (!NumberHasDecimals(result))
                    {
                        cellValue = int.Parse(result.ToString());
                    }
                    else
                    {
                        cellValue = Math.Round(result, 3);
                    }
                }
                else if (cell.CellValue != null)
                {
                    cellValue = cell.CellValue.InnerText;
                }
            return cellValue;
        }

        internal static bool NumberHasDecimals(decimal input)
        {
            var hasDecimals = true;

            if (input % 1 == 0)
            {
                hasDecimals = false;
            }
            return hasDecimals;
        }

        internal static bool CheckIfFormatIsDate(WorkbookPart workbookPart, Cell cell)
        {
            var isDate = false;

            var cellFormats = workbookPart.WorkbookStylesPart?.Stylesheet.CellFormats;

            var cellFormat = cellFormats?.Descendants<CellFormat>().ElementAt(Convert.ToInt32(cell.StyleIndex?.Value));

            var numberingFormats = workbookPart.WorkbookStylesPart?.Stylesheet.NumberingFormats;

            if (cellFormat?.NumberFormatId! != null)
            {
                var numberFormatId = cellFormat.NumberFormatId.Value;

                var numberingFormat = numberingFormats.Cast<NumberingFormat>().SingleOrDefault(f => f.NumberFormatId?.Value == numberFormatId);

                if (numberingFormat != null && numberingFormat.FormatCode?.Value != null && numberingFormat.FormatCode.Value.Contains("yy"))
                {
                    isDate = true;
                }
            }
            return isDate;
        }

        internal static bool CheckIfFormatIsNumber(WorkbookPart workbookPart, Cell cell)
        {
            var isNumber = false;

            var numberFormatIds = new List<uint> { 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 164, 166, 167, 168, 169, 171, 173, 197, 198, 199, 200, 201, 202, 203, 204, 205, 207 };

            var cellFormats = workbookPart.WorkbookStylesPart?.Stylesheet.CellFormats;

            var cellFormat = cellFormats?.Descendants<CellFormat>().ElementAt(Convert.ToInt32(cell.StyleIndex!.Value));

            if (cellFormat?.NumberFormatId! != null && numberFormatIds.Contains(cellFormat.NumberFormatId))
            {
                isNumber = true;
            }
            return isNumber;
        }

        internal static Dictionary<string, string>? GetSheetsForFilteredDefinedNames(WorkbookPart workbookPart, Dictionary<string, string> filteredDefinedNames)
        {
            var sheetsNameRelationshipIdPairs = new Dictionary<string, string>();

            const string nameAttribute = "name";

            const string idAttribute = "id";

            if (workbookPart.Workbook.Sheets != null)
            {
                foreach (var sheet in workbookPart.Workbook.Sheets)
                {
                    var sheetName = string.Empty;

                    var relationshipId = string.Empty;

                    foreach (var attr in sheet.GetAttributes())
                    {
                        switch (attr.LocalName)
                        {
                            case nameAttribute:
                                sheetName = attr.Value;
                                break;

                            case idAttribute:
                                relationshipId = attr.Value;
                                break;
                        }
                    }
                    foreach (var dn in filteredDefinedNames.Values.Where(dn => sheetName != null && dn.Contains(sheetName) & !sheetsNameRelationshipIdPairs.ContainsKey(sheetName)))
                    {
                        sheetsNameRelationshipIdPairs.Add(sheetName, relationshipId);
                    }
                }
            }
            else
            {
                return null;
            }
            return sheetsNameRelationshipIdPairs;
        }

        internal static WorksheetPart? GetSheetDataByRelationshipId(WorkbookPart workbookPart, string relationshipId)
        {
            WorksheetPart? worksheetPart = null;

            foreach (var wp in workbookPart.WorksheetParts)
            {
                var partRelationshipId = workbookPart.GetIdOfPart(wp);

                if (partRelationshipId == relationshipId)
                {
                    worksheetPart = wp;
                }
            }
            return worksheetPart;
        }

        internal static Cell? GetCell(WorksheetPart worksheetPart, string columnName, uint rowIndex, bool excludeHiddenRow, bool excludeHiddenColumn)
        {
            Cell? cell = null;

            var row = GetRow(worksheetPart, rowIndex);

            if (row == null || excludeHiddenRow && CheckIfCellInHiddenRow(row))
            {
                return cell;
            }
            cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference != null && c.CellReference.Value != null
                   && string.Equals(c.CellReference?.Value, columnName + rowIndex, StringComparison.CurrentCultureIgnoreCase));

            if (cell == null || (excludeHiddenColumn && CheckIfCellInHiddenColumn(worksheetPart, row, cell)))
            {
                return null;
            }
            return cell;

        }

        internal static bool CheckIfCellInHiddenRow(Row row)
        {
            var isHidden = row.Hidden! != null! && row.Hidden.Value;

            return isHidden;
        }

        internal static bool CheckIfCellInHiddenColumn(WorksheetPart worksheetPart, Row row, Cell cell)
        {
            var isHidden = false;
            Columns? columns = null;
            var hiddenColumnNames = new HashSet<string>();

            if (worksheetPart.Worksheet.GetFirstChild<Columns>() != null)
            {
                columns = worksheetPart.Worksheet.Elements<Columns>().First();

                foreach (var col in columns.Elements<Column>().Where(c => c.Hidden! != null! && c.Hidden != null! && c.Hidden.Value))
                {
                    for (uint min = col.Min!, max = col.Max!; min <= max; min++)
                    {
                        hiddenColumnNames.Add(GetColumnName(min));
                    }
                }
                var column = cell.CellReference?.Value?.Replace(row.RowIndex?.ToString()!, "");

                if (column != null && hiddenColumnNames.Contains(column))
                {
                    isHidden = true;
                }
            }
            return isHidden;
        }

        internal static string GetColumnName(uint columnNumber)
        {
            var columnName = "";

            while (columnNumber > 0)
            {
                var modulo = (columnNumber - 1) % 26;

                columnName = Convert.ToChar(65 + modulo) + columnName;

                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }

        internal static Row? GetRow(WorksheetPart worksheetPart, uint rowIndex)
        {
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            return sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex! == rowIndex);
        }

        internal static SharedStringItem? GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }

        internal static CultureInfo GetCultureWithCustomNumberFormat()
        {
            CultureInfo cultureInfo = new CultureInfo("en-US");
            NumberFormatInfo numberFormat = new NumberFormatInfo
            {
                NumberDecimalDigits = 3
            };
            cultureInfo.NumberFormat = numberFormat;

            return cultureInfo;
        }
    }
}
