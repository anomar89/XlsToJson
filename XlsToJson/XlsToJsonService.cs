using DocumentFormat.OpenXml;
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

            var filteredDefinedNames = new Dictionary<string, string>();

            var definedNameValuePairs = new Dictionary<string, string>();

            if (workbookPart == null || workbookPart.Workbook == null || workbookPart.Workbook.DefinedNames == null)
            {
                return bcsJson;
            }
            DefinedNames definedNames = workbookPart.Workbook.DefinedNames;

            filteredDefinedNames = FilterDefinedNamesWithRegex(definedNames, filters);

            if (filteredDefinedNames == null || filteredDefinedNames.Count == 0)
            {
                return bcsJson;
            }
            var filteredSheets = GetSheetsForFilteredDefinedNames(workbookPart, filteredDefinedNames);

            if (filteredSheets == null || filteredSheets.Count == 0)
            {
                return bcsJson;
            }
            foreach (var sheet in filteredSheets)
            {
                var sheetData = GetSheetDataByRelationshipId(workbookPart, sheet.Value);

                if (sheetData == null)
                {
                    continue;
                }
                foreach (var definedName in filteredDefinedNames.Where(fdn => fdn.Value.Contains(sheet.Key)))
                {
                    if (!definedName.Value.Contains('$'))
                    {
                        continue;
                    }
                    var columnName = definedName.Value.Split('$', '$')[1];

                    var rowIndex = uint.Parse(definedName.Value[(definedName.Value.LastIndexOf("$") + 1)..]);

                    Cell? cell = GetCell(sheetData, columnName, rowIndex, excludeHiddenRows, excludeHiddenColumns);

                    if (cell == null)
                    {
                        continue;
                    }
                    var cellValue = GetCellValue(workbookPart, cell);

                    if (definedName.Value.Contains(sheet.Key) && !string.IsNullOrEmpty(cellValue))
                    {
                        definedNameValuePairs.Add(definedName.Key, cellValue);
                    }
                }
            }
            if (definedNameValuePairs != null && definedNameValuePairs.Count > 0)
            {
                bcsJson = JsonConvert.SerializeObject(definedNameValuePairs);
            }
            return bcsJson;
        }

        internal static Dictionary<string, string> FilterDefinedNamesWithRegex(DefinedNames definedNames, Regex[] filters)
        {
            Dictionary<string, string> filteredDefinedNames = new Dictionary<string, string>();

            foreach (DefinedName definedName in definedNames)
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

        internal static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            string cellValue = string.Empty;

            if (cell.DataType != null)
            {
                if (cell.DataType == CellValues.SharedString)
                {
                    if (int.TryParse(cell.InnerText, out int id))
                    {
                        SharedStringItem? item = GetSharedStringItemById(workbookPart, id);

                        if (item == null)
                        {
                            return cellValue;
                        }
                        if (item.Text != null)
                        {
                            cellValue = item.Text.Text;
                        }
                        else if (item.InnerText != null)
                        {
                            cellValue = item.InnerText;
                        }
                        else if (item.InnerXml != null)
                        {
                            cellValue = item.InnerXml;
                        }
                    }
                }
            }
            else if (cell.CellValue != null)
            {
                cellValue = cell.CellValue.InnerText;
            }
            return cellValue;
        }

        internal static Dictionary<string, string>? GetSheetsForFilteredDefinedNames(WorkbookPart workbookPart, Dictionary<string, string> filteredDefinedNames)
        {
            var sheetsNameRelationshipIdPairs = new Dictionary<string, string>();

            var nameAttribute = "name";

            var idAttribute = "id";

            if (workbookPart.Workbook.Sheets != null)
            {
                foreach (var sheet in workbookPart.Workbook.Sheets)
                {
                    var sheetname = string.Empty;

                    var relationshipId = string.Empty;

                    foreach (OpenXmlAttribute attr in sheet.GetAttributes())
                    {
                        if (attr.LocalName == nameAttribute)
                        {
                            sheetname = attr.Value;
                        }
                        if (attr.LocalName == idAttribute)
                        {
                            relationshipId = attr.Value;
                        }
                    }
                    foreach (var dn in filteredDefinedNames.Values)
                    {
                        if (dn.Contains(sheetname) & !sheetsNameRelationshipIdPairs.ContainsKey(sheetname))
                        {
                            sheetsNameRelationshipIdPairs.Add(sheetname, relationshipId);
                        }
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
                string partRelationshipId = workbookPart.GetIdOfPart(wp);

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

            Row? row = GetRow(worksheetPart, rowIndex);

            if (row == null || excludeHiddenRow && CheckIfCellInHiddenRow(row))
            {
                return cell;
            }
            cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference != null && c.CellReference.Value != null && string.Equals(c.CellReference?.Value, columnName + rowIndex, StringComparison.CurrentCultureIgnoreCase));
            
            if (cell == null || (cell != null && excludeHiddenColumn && CheckIfCellInHiddenColumn(worksheetPart, row, cell)))
            {
                return null;
            }
            return cell;

        }

        internal static bool CheckIfCellInHiddenRow(Row row)
        {
            bool isHidden = false;

            if (row.Hidden != null && row.Hidden.Value == true)
            {
               isHidden = true;
            }
            return isHidden;
        }

        internal static bool CheckIfCellInHiddenColumn(WorksheetPart worksheetPart, Row row, Cell cell)
        {
            bool isHidden = false;

            var columns = worksheetPart.Worksheet.Elements<Columns>().First();

            var hiddenColumnNames = new HashSet<string>();

            foreach (var col in columns.Elements<Column>().Where(c => c.Hidden != null && c.Hidden.Value))
            {
                for (uint min = col.Min, max = col.Max; min <= max; min++)
                {
                    hiddenColumnNames.Add(GetColumnName(min));
                }
            }
            var column = cell.CellReference?.Value?.Replace(row.RowIndex?.ToString(), "");

            if (column != null && hiddenColumnNames.Contains(column))
            {
               isHidden = true;
            }
            return isHidden;
        }

        internal static string GetColumnName(uint columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                uint modulo = (columnNumber - 1) % 26;

                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;

                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnName;
        }

        internal static Row? GetRow(WorksheetPart worksheetPart, uint rowIndex)
        {
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            return sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        }

        internal static SharedStringItem? GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }
    }
}
