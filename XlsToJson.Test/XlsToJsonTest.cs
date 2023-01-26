namespace XlsToJson.Test
{
    [TestClass]
    public class XlsToJsonTest
    {
        private const string path = @"assets/sample.xlsm";

        [TestMethod]
        public void ConvertXlsToJson_ByFilePath_Success()
        {
            var json = XlsToJson.ConvertXlsToJson(path);

            Assert.IsNotNull(json);
            Assert.IsTrue(json.HasValues);

            Validation.ValidateDateTime(json);
        }

        [TestMethod]
        public void ConvertXlsToJson_ByMemoryStream_Success()
        {
            byte[] fileContent = File.ReadAllBytes(path);

            using MemoryStream ms = new(fileContent);
            var json = XlsToJson.ConvertXlsToJson(ms);

            Assert.IsNotNull(json);
            Assert.IsTrue(json.HasValues);

            Validation.ValidateDateTime(json);
        }
    }
}