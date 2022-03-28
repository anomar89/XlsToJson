using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace XlsToJson
{
    internal static class XlsToJsonService
    {
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

            if (workbookPart.Workbook.Sheets != null)
            {
                foreach (var sheet in workbookPart.Workbook.Sheets)
                {
                    var sheetname = string.Empty;

                    var relationshipId = string.Empty;

                    foreach (OpenXmlAttribute attr in sheet.GetAttributes())
                    {
                        if (attr.LocalName == "name")
                        {
                            sheetname = attr.Value;
                        }
                        if (attr.LocalName == "id")
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

        internal static SheetData? GetSheetDataByRelationshipId(WorkbookPart workbookPart, string relationshipId)
        {
            SheetData? sheetData = null;

            foreach (var w in workbookPart.WorksheetParts)
            {
                string partRelationshipId = workbookPart.GetIdOfPart(w);

                if (partRelationshipId == relationshipId)
                {
                    sheetData = w.Worksheet.Elements<SheetData>().First();
                }
            }
            return sheetData;
        }

        internal static Cell? GetCell(SheetData worksheet, string columnName, uint rowIndex)
        {
            Row? row = GetRow(worksheet, rowIndex);

            if (row == null)
            {
                return null;
            }
            if (row.Hidden != null && row.Hidden.Value == true)
            {
                return null;
            }
            return row.Elements<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, columnName + rowIndex, StringComparison.CurrentCultureIgnoreCase));

        }

        internal static Row? GetRow(SheetData worksheet, uint rowIndex)
        {
            return worksheet.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        }

        internal static SharedStringItem? GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart?.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }
    }
}
