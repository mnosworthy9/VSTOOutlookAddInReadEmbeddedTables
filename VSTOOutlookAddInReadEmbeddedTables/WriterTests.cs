using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace VSTOOutlookAddInReadEmbeddedTables
{
    [TestClass]
    public class WriterTests
    {
        [TestMethod]
        public void WriteToExcel()
        {
            // Arrange 
            string[] values = { "value", "test" };

            // Act
            ExcelWriter.WriteToExcel(values);

            // Assert

        }
    }
}
