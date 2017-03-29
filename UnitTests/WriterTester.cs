using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;
using XLOC.Book;
using XLOC.Utility;
using XLOC;

namespace ExcelReaderUnitTestProject
{
    [TestClass]
    public class WriterTester
    {
        [TestMethod]
        public void WriteEmptyBook()
        {
            var book = new XlBook();
            book.AddSheet("test1");
            book.Sheets.First().AddCell(12.3, "B2", XlContentType.Double);
            book.Sheets.First().AddCell(DateTime.Now, "C3", XlContentType.Date);
            book.Sheets.First().AddCell("asdasd", "D4", XlContentType.SharedString);
            var err = XlConverter.FromBook(book).SaveToStream(new System.IO.MemoryStream());
            if (err.Count() > 0)
                Assert.Fail("Ошибка сохранения:\n{0}", string.Join("\n", err.Select(x => x.Description)));
        }
    }
}