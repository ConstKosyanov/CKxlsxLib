using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using XLOC.Book;
using XLOC.Utility;

namespace ExcelReaderUnitTestProject
{
    [TestClass]
    public class XlSheetTester
    {
        XlSheet xs;

        [TestInitialize]
        public void Initialize()
        {
            var xl = new XlBook();
            xl.AddSheet("testList");
            xs = xl.Sheets.First();
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void NameValidation() => xs.Name = "test:";

        [TestMethod, ExpectedException(typeof(IndexOutOfRangeException))]
        public void FalseCellIndex() => xs.Get(0, 0);

        [TestMethod]
        public void NameSet()
        {
            xs.Name = "test";
            Assert.AreEqual("test", xs.Name);
        }

        [TestMethod]
        public void AddCell()
        {
            xs.AddCell(1, 1, 1, XlContentType.Integer);
            Assert.AreEqual(1, xs.Cells.Count());
        }
    }
}