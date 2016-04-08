using qXlsxLib;
using qXlsxLib.Excel;
using qXlsxLib.Writer;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using qXlsxLib.Utility;

namespace ExcelReaderUnitTestProject
{
    [TestClass]
    public class WriterTester
    {
        [TestMethod]
        public void WriteEmptyBook()
        {
            var book = new xlBook();
            book.AddSheet("test1");
            book.Sheets.First().AddCell(12.3, "B2", xlContentType.Double);
            book.Sheets.First().AddCell(DateTime.Now, "C3", xlContentType.Date);
            book.Sheets.First().AddCell("asdasd", "D4", xlContentType.SharedString);
            var err = xlWriter.Create(book).SaveToStream(new System.IO.MemoryStream());
            if (err.Count() > 0)
                Assert.Fail("Ошибка сохранения:\n{0}", string.Join("\n", err.Select(x => x.Description)));
        }
    }
}