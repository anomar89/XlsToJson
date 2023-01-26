namespace XlsToJson.Test
{
    internal class Validation
    {
        private static readonly DateTime _defaultDateValue = new(2023, 01, 22);

        /// <summary>
        /// Check if input dates are valid
        /// Default date is set to 22/01/2023
        /// </summary>
        /// <param name="json"></param>
        /// <exception cref="Exception"></exception>
        internal static void ValidateDateTime(JObject json)
        {
            var inputDates = json.Properties().Where(i => i.Name.StartsWith("date")).Select(i => i).ToList();

            if (inputDates.Any())
            {
                foreach (var date in inputDates)
                {
                    _ = DateTime.TryParse(date.Value.ToString(), out var dateValue);

                    if (dateValue.Date == default(DateTime).Date)
                    {
                        throw new Exception($"{date.Value} is not a valid date.");
                    }

                    if (dateValue.Date == _defaultDateValue.Date)
                    {
                        continue;
                    }
                    else
                    {
                        throw new Exception($"Input date is not equal to default date of `{_defaultDateValue.ToShortDateString()}`.");
                    }
                }
            }
            else
            {
                throw new Exception("No dates were found. Please check input values!");
            }
        }
    }
}