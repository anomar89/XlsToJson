using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace XlsToJson
{
    internal static class XlsToJsonService
    {
        internal static string? ProcessDocument(WorkbookPart workbookPart, Regex[]? filters, bool excludeHiddenRows, bool excludeHiddenColumns)
        {
            var bcsJson = string.Empty;

            var definedNameValuePairs = new Dictionary<string, string>();

            if (workbookPart?.Workbook?.DefinedNames == null)
            {
                return bcsJson;
            }
            var definedNames = workbookPart.Workbook.DefinedNames;

            var filteredDefinedNames = FilterDefinedNamesWithRegex(definedNames, filters);

            if (filteredDefinedNames.Count == 0)
            {
                return bcsJson;
            }
            var filteredSheets = GetSheetsForFilteredDefinedNames(workbookPart, filteredDefinedNames);

            if (filteredSheets == null || filteredSheets.Count == 0)
            {
                return bcsJson;
            }
            foreach (var (key, value) in filteredSheets)
            {
                var sheetData = GetSheetDataByRelationshipId(workbookPart, value);

                if (sheetData == null)
                {
                    continue;
                }
                foreach (var (s, value1) in filteredDefinedNames.Where(fdn => fdn.Value.Contains(key)))
                {
                    if (!value1.Contains('$'))
                    {
                        continue;
                    }
                    var columnName = value1.Split('$', '$')[1];

                    var rowIndex = uint.Parse(value1[(value1.LastIndexOf("$", StringComparison.Ordinal) + 1)..]);

                    var cell = GetCell(sheetData, columnName, rowIndex, excludeHiddenRows, excludeHiddenColumns);

                    if (cell == null)
                    {
                        continue;
                    }
                    var cellValue = GetCellValue(workbookPart, cell);

                    if (value1.Contains(key) && !string.IsNullOrEmpty(cellValue))
                    {
                        definedNameValuePairs.Add(s, cellValue);
                    }
                }
            }
            if (definedNameValuePairs.Count > 0)
            {
                bcsJson = JsonConvert.SerializeObject(definedNameValuePairs);
            }
            return bcsJson;
        }

        internal static Dictionary<string, string> FilterDefinedNamesWithRegex(DefinedNames definedNames, Regex[]? filters)
        {
            var filteredDefinedNames = new Dictionary<string, string>();

            foreach (var openXmlElement in definedNames)
            {
                var definedName = (DefinedName)openXmlElement;

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

        internal static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            var cellValue = string.Empty;

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
            else if (cell.StyleIndex != null && cell.CellValue != null && CheckIfFormatIsDate(workbookPart, cell))
            {
                cellValue = DateTime.FromOADate(Convert.ToDouble(cell.CellValue.Text)).ToShortDateString();
            }
            else if (cell.CellValue != null)
            {
                cellValue = cell.CellValue.InnerText;
            }
            return cellValue;
        }

        internal static bool CheckIfFormatIsDate(WorkbookPart workbookPart, Cell cell)
        {
            var isDate = false;

            var dateFormatCodes = new List<uint> { 166, 175, 179, 182, 183, 189, 193, 194, 196 };

            var cellFormats = workbookPart.WorkbookStylesPart?.Stylesheet.CellFormats;

            var cellFormat = cellFormats?.Descendants<CellFormat>().ElementAt(Convert.ToInt32(cell.StyleIndex.Value));

            if (cellFormat.NumberFormatId != null! && dateFormatCodes.Contains(cellFormat.NumberFormatId))
            {
                isDate = true;
            }
            return isDate;
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

            var columns = worksheetPart.Worksheet.Elements<Columns>().First();

            var hiddenColumnNames = new HashSet<string>();

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
    }
}
