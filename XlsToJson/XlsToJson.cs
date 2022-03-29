using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace XlsToJson
{
    public static class XlsToJson
    {
        /// <summary>
        /// Extracts the list of defined names and their associated values from a XLSM file based on the provided regex filtering and returns a JSON with them
        /// </summary>
        /// <param name="filePath">The path to the XLSM file that you wish to process</param>
        /// <param name="filters">An array of regular expressions that can be used for filtering the defined names. It is optional and, if left null, all the defined names will be included in the JSON object with the exception of the ones without associated values</param>
        /// <param name="excludeHiddenRows">It is a Boolean that allows to control whether to include the cells from hidden rows in the JSON result. It is optional and true by default</param>
        /// <param name="excludeHiddenColumns">It is a Boolean that allows to control whether to include the cells from hidden columns in the JSON result. It is optional and true by default</param>
        ///<returns>It returns a nullable string containing the filtered defined names with their associated values in the JSON format</returns>
        public static string ConvertXlsToJson(string filePath, Regex[]? filters = null, bool excludeHiddenRows = true, bool excludeHiddenColumns = true)
        {
            try
            {
                var fieldsValuesPairs = new Dictionary<string, string>();

                var filteredDefinedNames = new Dictionary<string, string>();

                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;

                    if (workbookPart == null || workbookPart.Workbook == null || workbookPart.Workbook.DefinedNames == null)
                    {
                        return string.Empty;
                    }
                    DefinedNames definedNames = workbookPart.Workbook.DefinedNames;

                    filteredDefinedNames = XlsToJsonService.FilterDefinedNamesWithRegex(definedNames, filters);

                    var filteredSheets = XlsToJsonService.GetSheetsForFilteredDefinedNames(workbookPart, filteredDefinedNames);

                    if (filteredSheets == null || filteredSheets.Count == 0)
                    {
                        return string.Empty;
                    }
                    foreach (var sheet in filteredSheets)
                    {
                        var sheetData = XlsToJsonService.GetSheetDataByRelationshipId(workbookPart, sheet.Value);

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

                            Cell? cell = XlsToJsonService.GetCell(sheetData, columnName, rowIndex, excludeHiddenRows, excludeHiddenColumns);

                            if (cell == null)
                            {
                                continue;
                            }
                            var cellValue = XlsToJsonService.GetCellValue(workbookPart, cell);

                            if (definedName.Value.Contains(sheet.Key) && !string.IsNullOrEmpty(cellValue))
                            {
                                fieldsValuesPairs.Add(definedName.Key, cellValue);
                            }
                        }
                    }
                }
                var bcsJson = JsonConvert.SerializeObject(fieldsValuesPairs);

                return bcsJson;
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.Message);

                return string.Empty;
            }
        }
    }
}
