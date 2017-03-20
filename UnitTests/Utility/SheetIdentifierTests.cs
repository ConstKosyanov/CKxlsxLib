using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace XLOC.Utility.Tests
{
    [TestClass]
    public class SheetIdentifierTests
    {
        [TestMethod]
        public void SheetIdentifierTest()
        {
            var res = new SheetIdentifier(new DocumentFormat.OpenXml.Spreadsheet.Sheet { Id = "1", Name = "test" });
            Assert.AreEqual("1", res.Id);
            Assert.AreEqual("test", res.Name);
        }
    }
}