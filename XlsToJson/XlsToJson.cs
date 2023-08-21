using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;

namespace XlsToJson
{
    public static class XlsToJson
    {
        public static JObject ConvertXlsToJson(string filePath, out string errorMessage, Regex[]? filters = null, bool excludeHiddenRows = true, bool excludeHiddenColumns = true)
        {
            return ConvertXlsToJsonObject(() => SpreadsheetDocument.Open(filePath, false), out errorMessage, filters, excludeHiddenRows, excludeHiddenColumns);
        }

        public static JObject ConvertXlsToJson(MemoryStream fileContents, out string errorMessage, Regex[]? filters = null, bool excludeHiddenRows = true, bool excludeHiddenColumns = true)
        {
            return ConvertXlsToJsonObject(() => SpreadsheetDocument.Open(fileContents, false), out errorMessage, filters, excludeHiddenRows, excludeHiddenColumns);
        }

        private static JObject ConvertXlsToJsonObject(Func<SpreadsheetDocument> getDocument, out string errorMessage, Regex[]? filters = null, bool excludeHiddenRows = true, bool excludeHiddenColumns = true)
        {
            Thread.CurrentThread.CurrentCulture = XlsToJsonService.GetCultureWithCustomNumberFormat();

            errorMessage = string.Empty;
            var bcsJson = new JObject();

            try
            {
                using var spreadsheetDocument = getDocument();

                bcsJson = ProcessDocument(spreadsheetDocument, filters, excludeHiddenRows, excludeHiddenColumns);
            }
            catch (Exception ex)
            {
                errorMessage = $"The parsing of the BCS failed with the error: {ex.Message} Details: {ex.StackTrace}";
            }
            return bcsJson;
        }

        private static JObject ProcessDocument(SpreadsheetDocument spreadsheetDocument, Regex[]? filters, bool excludeHiddenRows, bool excludeHiddenColumns)
        {
            var workbookPart = spreadsheetDocument.WorkbookPart;

            if (workbookPart == null)
            {
                return new JObject();
            }
            return XlsToJsonService.ProcessDocument(workbookPart, filters, excludeHiddenRows, excludeHiddenColumns);
        }

      
    }
}
