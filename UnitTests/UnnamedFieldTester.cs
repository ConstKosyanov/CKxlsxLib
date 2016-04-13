using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Linq;
using XLOC.Book;
using XLOC.Utility;
using XLOC.Writer;

namespace ExcelReaderUnitTestProject
{
    [TestClass]
    public class UnnamedFieldTester
    {
        class MyClass : IxlCompatible
        {
            [xlField(xlContentType.Integer)]
            public int Field { get; set; }
        }

        [TestMethod]
        public void TestMethod1()
        {
            var book = new xlBook();
            var sh = book.AddSheet("sheet1");
            sh.AddCell("Field", "A1", xlContentType.SharedString);
            sh.AddCell("1", "A2", xlContentType.SharedString);
            sh.AddCell(1, "A3", xlContentType.Double);
            sh.AddCell(1, "A4", xlContentType.Integer);
            sh.AddCell(1, "A5", xlContentType.Void);

            var memstream = new MemoryStream();
            var err = xlWriter.Create(book).SaveToStream(memstream);
            Assert.IsFalse(err.Any(), string.Join("\n", err.Select(x => x.Description)));

            var result = XLOC.XlConverter.FromStream(memstream).ReadToArray<MyClass>();
            Assert.AreEqual(4, result.Count());
            Assert.IsTrue(result.All(x => x.Field == 1));
        }
    }
}