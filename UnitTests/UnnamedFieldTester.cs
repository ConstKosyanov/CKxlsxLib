using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLOC.Book;
using XLOC.Utility;
using XLOC.Writer;

namespace ExcelReaderUnitTestProject
{
    [TestClass]
    public class UnnamedFieldTester
    {
        class MyClass
        {
            [XlField(XlContentType.Integer)]
            public int Field { get; set; }
        }

        [TestMethod]
        public void TestMethod1()
        {
            var book = new XlBook();
            var sh = book.AddSheet("sheet1");
            sh.AddCell("Field", "A1", XlContentType.SharedString);
            sh.AddCell("1", "A2", XlContentType.SharedString);
            sh.AddCell(1, "A3", XlContentType.Double);
            sh.AddCell(1, "A4", XlContentType.Integer);
            sh.AddCell(1, "A5", XlContentType.Void);

            var memstream = new MemoryStream();
            var err = XlWriter.Create(book).SaveToStream(memstream);
            Assert.IsFalse(err.Any(), string.Join("\n", err.Select(x => x.Description)));

            var result = XLOC.XlConverter.FromStream(memstream).ReadToArray<MyClass>();
            Assert.AreEqual(4, result.Count());
            Assert.IsTrue(result.All(x => x.Field == 1));
        }
    }
}