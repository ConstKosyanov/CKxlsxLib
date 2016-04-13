using XLOC;
using XLOC.Book;
using XLOC.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using XLOC.Utility;

namespace ExcelReaderUnitTestProject
{
    [TestClass]
    public class xlCellTester
    {
        xlBook xl = new xlBook();
        xlSheet xs;

        [TestInitialize]
        public void TestInitialize()
        {
            xl = new xlBook();
            xs = xl.AddSheet("test");
        }

        [TestMethod]
        public void Constructor0()
        {
            var cell = xs.AddCell(1, 1, 1, xlContentType.Integer);
            Assert.IsNotNull(cell);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentOutOfRangeException))]
        public void Constructor1()
        {
            var cell = xs.AddCell(DateTime.Now, 0, 0, xlContentType.Date);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void Reference0()
        {
            var cell = xs.AddCell(DateTime.Now, 1, 1, xlContentType.Date);
            cell.Reference = "asdasd a";
        }

        [TestMethod]
        public void Reference1()
        {
            var cell = xs.AddCell(1, 1, 1, xlContentType.Integer);
            cell.Reference = "A1";
            Assert.AreEqual("A1", cell.Reference);
            Assert.AreEqual(1, cell.Row);
            Assert.AreEqual(1, cell.Col);
            cell.Reference = "B2";
            Assert.AreEqual("B2", cell.Reference);
            Assert.AreEqual(2, cell.Row);
            Assert.AreEqual(2, cell.Col);
            cell.Reference = "BB48";
            Assert.AreEqual("BB48", cell.Reference);
            Assert.AreEqual(48, cell.Row);
            Assert.AreEqual(54, cell.Col);
            cell.Reference = "AR1";
            Assert.AreEqual("AR1", cell.Reference);
            Assert.AreEqual(1, cell.Row);
            Assert.AreEqual(44, cell.Col);
            cell.Reference = "XFD010";
            Assert.AreEqual("XFD10", cell.Reference);
            Assert.AreEqual(10, cell.Row);
            Assert.AreEqual(16384, cell.Col);
        }

        [TestMethod]
        public void Reference()
        {
            var cell = xs.AddCell(DateTime.Now, 1, 1, xlContentType.Date);
            Assert.AreEqual("A1", cell.Reference);
            cell.Col = 15;
            cell.Row = 2;
            Assert.AreEqual("O2", cell.Reference);
            cell.Col = 78;
            cell.Row = 3;
            Assert.AreEqual("BZ3", cell.Reference);
        }
    }
}
