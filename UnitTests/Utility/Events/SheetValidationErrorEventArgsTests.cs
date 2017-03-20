using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

namespace XLOC.Utility.Events.Tests
{
    [TestClass]
    public class SheetValidationErrorEventArgsTests
    {
        [TestMethod]
        public void SheetValidationErrorEventArgsTest()
        {
            var res = new SheetValidationErrorEventArgs(new[] { "a", "b", "c" }, new[] { new Exception("a"), new Exception("b") }, new SheetIdentifier(new DocumentFormat.OpenXml.Spreadsheet.Sheet { Id = "1", Name = "test" }));

            string[] tmp = new[] { "a", "b", "c" };
            Assert.AreEqual(3, res.MissingFields.Count());
            for (int i = 0; i < res.MissingFields.Count(); i++)
                Assert.AreEqual(tmp[i], res.MissingFields.ToList()[i]);

            Assert.AreEqual(2, res.Exceptions.Count());
            Assert.AreEqual("a", res.Exceptions.First().Message);
            Assert.AreEqual("b", res.Exceptions.Last().Message);

            Assert.AreEqual("1", res.Sheet.Id);
            Assert.AreEqual("test", res.Sheet.Name);
        }
    }
}