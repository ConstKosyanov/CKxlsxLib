using qXlsxLib;
using qXlsxLib.Book;
using qXlsxLib.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using qXlsxLib.Utility;

namespace ExcelReaderUnitTestProject
{
    [TestClass]
    public class xlSheetTester
    {
        xlSheet xs;

        [TestInitialize]
        public void Initialize()
        {
            var xl = new xlBook();
            xl.AddSheet("testList");
            xs = xl.Sheets.First();
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void NameValidation()
        {
            xs.Name = "test:";
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void FalseCellIndex()
        {
            var tmp = xs.Get(0, 0);
        }

        [TestMethod]
        public void NameSet()
        {
            xs.Name = "test";
            Assert.AreEqual("test", xs.Name);
        }

        [TestMethod]
        public void AddCell()
        {
            xs.AddCell(1,1,1, xlContentType.Integer);
            Assert.AreEqual(1, xs.Cells.Count());
        }
    }
}
