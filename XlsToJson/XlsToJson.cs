using DocumentFormat.OpenXml.Packaging;
using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace XlsToJson
{
    public static class XlsToJson
    {
        /// <summary>
        /// Extracts the filtered list of defined names and their associated values from a XLS file stored on the disk and returns a JSON result
        /// </summary>
        /// <param name="filePath">A path to the XLS file to process</param>
        /// <param name="filters">A collection of regular expressions used for filtering the defined names. It is optional and, if left null, all the defined names will be included in the JSON result except the ones without associated values</param>
        /// <param name="excludeHiddenRows">A flag that excludes cells located on hidden rows from the JSON result. It is optional and true by default</param>
        /// <param name="excludeHiddenColumns">A flag that excludes cells located on hidden columns from the JSON result. It is optional and true by default</param>
        ///<returns>It returns a nullable string containing the filtered defined names with their associated values in the JSON format</returns>
        public static string? ConvertXlsToJson(string filePath, Regex[]? filters = null, bool excludeHiddenRows = true, bool excludeHiddenColumns = true)
        {
            var bcsJson = string.Empty;
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;

                    if (workbookPart == null)
                    {
                        return bcsJson;
                    }
                    bcsJson = XlsToJsonService.ProcessDocument(workbookPart, filters, excludeHiddenRows, excludeHiddenColumns);
                }
                return bcsJson;
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.Message);

                return bcsJson;
            }
        }

        /// <summary>
        /// Extracts the filtered list of defined names and their associated values from a XLS file stored in memory and returns a JSON result
        /// </summary>
        /// <param name="fileContents">A memory stream containing the XLS file to process</param>
        /// <param name="filters">A collection of regular expressions used for filtering the defined names. It is optional and, if left null, all the defined names will be included in the JSON result except the ones without associated values</param>
        /// <param name="excludeHiddenRows">A flag that excludes cells located on hidden rows from the JSON result. It is optional and true by default</param>
        /// <param name="excludeHiddenColumns">A flag that excludes cells located on hidden columns from the JSON result. It is optional and true by default</param>
        ///<returns>It returns a nullable string containing the filtered defined names with their associated values in the JSON format</returns>
        public static string? ConvertXlsToJson(MemoryStream fileContents, Regex[]? filters = null, bool excludeHiddenRows = true, bool excludeHiddenColumns = true)
        {
            var bcsJson = string.Empty;
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileContents, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;

                    if (workbookPart == null)
                    {
                        return bcsJson;
                    }
                    bcsJson = XlsToJsonService.ProcessDocument(workbookPart, filters, excludeHiddenRows, excludeHiddenColumns);
                }
                return bcsJson;
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.Message);

                return bcsJson;
            }

        }
    }
}
