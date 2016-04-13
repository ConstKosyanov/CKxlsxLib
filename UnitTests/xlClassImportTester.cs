using qXlsxLib;
using qXlsxLib.Excel;
using qXlsxLib.Reader;
using qXlsxLib.Writer;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using qXlsxLib.Utility;

namespace ExcelReaderUnitTestProject
{
    [TestClass]
    public class xlClassImportTester
    {
        static string path = string.Format(@"{0}\{1}", Path.Combine(Environment.CurrentDirectory), "test2.xlsx");

        TestExcelClass[] data = new TestExcelClass[]
        {
            new TestExcelClass() { intProperty1 = 1 , SomeDate = DateTime.Now, SomeString = "asdasd"},
            new TestExcelClass() { intProperty1 = 2 , SomeDate = DateTime.Now, SomeString = "aafgf"},
            new TestExcelClass() { intProperty1 = 3 , SomeDate = DateTime.Now, SomeString = "xdfe"},
            new TestExcelClass() { intProperty1 = 4 , SomeDate = DateTime.Now, SomeString = "dfdr"},
            new TestExcelClass() { intProperty1 = 5 , SomeDate = DateTime.Now, SomeString = "ghdg"},
            new TestExcelClass() { intProperty1 = 7 , SomeDate = DateTime.Now, SomeString = "dfg"},
            new TestExcelClass() { intProperty1 = 9 , SomeDate = DateTime.Now, SomeString = "dfgag"},
            new TestExcelClass() { intProperty1 = 10, SomeDate = DateTime.Now, SomeString = "sdfsw"},
        };

        [ClassInitialize]
        public static void Initialize(TestContext ctx)
        {
            File.Delete(path);
        }

        [TestMethod]
        public void Write()
        {
            var err = xlWriter.Create(data).SaveToFile(path);
            if (err.Count() > 0)
                Assert.Fail("Ошибка сохранения:\n{0}", string.Join("\n", err.Select(x => x.Description)));
        }

        [TestMethod]
        public void Read()
        {
            Write();
            var readedData = qXlsx.FromFile(path).ReadToEnumerable<TestExcelClass>().ToArray();
            Assert.AreEqual(data.Count(), readedData.Count(), "Количество загруженных строк не совпадает");
            for (int i = 0; i < data.Count(); i++)
            {
                Assert.AreEqual(data[i].intProperty1, readedData[i].intProperty1, "Поля заполены не верно");
                Assert.AreEqual(data[i].intProperty2, readedData[i].intProperty2, "Поля заполены не верно");
                Assert.AreEqual(data[i].SomeDate.ToShortDateString(), readedData[i].SomeDate.ToShortDateString(), "Поля заполены не верно");
                Assert.AreEqual(data[i].SomeString, readedData[i].SomeString, "Поля заполены не верно");
            }
        }

        [TestMethod]
        public void ReadToArrayWithoutNullableColumns()
        {
            var book = new xlBook();
            var sh = book.AddSheet("sheet1");

            #region Captions
            //=================================================
            sh.AddCell("Поле 1", "A1", xlContentType.SharedString);
            sh.AddCell("Какая-то дата", "B1", xlContentType.SharedString);
            sh.AddCell("Мультизагаловок2", "C1", xlContentType.SharedString);
            sh.AddCell("дробь", "E1", xlContentType.SharedString);
            //=================================================
            #endregion

            #region Data
            //=================================================
            sh.AddCell(1, "A2", xlContentType.Integer);
            sh.AddCell(DateTime.Now, "B2", xlContentType.Date);
            sh.AddCell("Какая-то строка", "C2", xlContentType.SharedString);
            sh.AddCell("0.15", "E2", xlContentType.Double);

            sh.AddCell(2, "A3", xlContentType.Integer);
            sh.AddCell(DateTime.Now, "B3", xlContentType.Date);
            sh.AddCell("Какая-то строка", "C3", xlContentType.SharedString);
            sh.AddCell("0.25", "E3", xlContentType.Double);
            //=================================================
            #endregion
            var memstream = new MemoryStream();
            xlWriter.Create(book).SaveToStream(memstream);

            TestExcelClass[] data = qXlsx.FromStream(memstream).ReadToEnumerable<TestExcelClass>().ToArray();
            Assert.AreEqual(2, data.Count());
            Assert.IsTrue(data.All(x => !x.intProperty2.HasValue));
        }

        [TestMethod]
        public void ReadToArrayWithNullableColumns()
        {
            var book = new xlBook();
            var sh = book.AddSheet("sheet1");

            #region Captions
            //=================================================
            sh.AddCell("Поле 1", "A1", xlContentType.SharedString);
            sh.AddCell("Какая-то дата", "B1", xlContentType.SharedString);
            sh.AddCell("Мультизагаловок2", "C1", xlContentType.SharedString);
            sh.AddCell("дробь", "E1", xlContentType.SharedString);
            sh.AddCell("Поле 3", "F1", xlContentType.SharedString);
            //=================================================
            #endregion

            #region Data
            //=================================================
            sh.AddCell(1, "A2", xlContentType.SharedString);
            sh.AddCell(DateTime.Now, "B2", xlContentType.Date);
            sh.AddCell("Какая-то строка", "C2", xlContentType.SharedString);
            sh.AddCell("0.15", "E2", xlContentType.Double);
            sh.AddCell("", "F2", xlContentType.SharedString);

            sh.AddCell(2, "A3", xlContentType.Integer);
            sh.AddCell(DateTime.Now, "B3", xlContentType.Date);
            sh.AddCell("Какая-то строка", "C3", xlContentType.SharedString);
            sh.AddCell("0.25", "E3", xlContentType.Double);
            sh.AddCell("", "F3", xlContentType.SharedString);
            //=================================================
            #endregion

            var memstream = new MemoryStream();
            xlWriter.Create(book).SaveToStream(memstream);

            TestExcelClass[] data = qXlsx.FromStream(memstream, new qXlsxConfiguration { CellReadingErrorEvent = (s, e) => { throw new Exception(e.Exception.Message); } }).ReadToArray<TestExcelClass>();
            Assert.AreEqual(2, data.Count());
            Assert.IsTrue(data.All(x => !x.intProperty3.HasValue));
        }
        
        public void MultiCaptionTest()
        {
            var book = new xlBook();
            var sh = book.AddSheet("sheet1");
            sh.AddCell("Поле 1", "A1", xlContentType.SharedString);
            sh.AddCell("Какая-то дата", "B1", xlContentType.SharedString);
            sh.AddCell("Мультизагаловок1", "C1", xlContentType.SharedString);
            sh.AddCell("Мультизагаловок2", "D1", xlContentType.SharedString);
            sh.AddCell("дробь", "E1", xlContentType.SharedString);
            sh.AddCell(1, "A2", xlContentType.Integer);
            sh.AddCell(DateTime.Now, "B2", xlContentType.Date);
            sh.AddCell("Какая-то строка", "C2", xlContentType.SharedString);
            sh.AddCell("Какая-то строка", "D2", xlContentType.SharedString);
            sh.AddCell("0.15", "E2", xlContentType.Double);
            sh.AddCell(2, "A3", xlContentType.Integer);
            sh.AddCell(DateTime.Now, "B3", xlContentType.Date);
            sh.AddCell("Какая-то строка", "C3", xlContentType.SharedString);
            sh.AddCell("Какая-то строка", "D3", xlContentType.SharedString);
            sh.AddCell("0.25", "E3", xlContentType.Double);
            var memstream = new MemoryStream();
            xlWriter.Create(book).SaveToStream(memstream);

            var isValid = true;
            TestExcelClass[] data = qXlsx.FromStream(memstream, new qXlsxConfiguration { ValidationFailureEvent = (s, e) => isValid = false }).ReadToEnumerable<TestExcelClass>().ToArray();
            Assert.AreEqual(2, data.Count());
            Assert.IsTrue(data.All(x => !x.intProperty2.HasValue));
            Assert.IsFalse(isValid);
        }

        [TestMethod]
        public void EventTest()
        {
            var book = new xlBook();
            var sh = book.AddSheet("sheet1");
            sh.AddCell("Какая-то дата", "B1", xlContentType.SharedString);
            sh.AddCell("Мультизагаловок1", "C1", xlContentType.SharedString);
            sh.AddCell("Мультизагаловок2", "D1", xlContentType.SharedString);
            sh.AddCell("дробь", "E1", xlContentType.SharedString);
            sh.AddCell(1, "A2", xlContentType.Integer);
            sh.AddCell(DateTime.Now, "B2", xlContentType.Date);
            sh.AddCell("Какая-то строка", "C2", xlContentType.SharedString);
            sh.AddCell("Какая-то строка", "D2", xlContentType.SharedString);
            sh.AddCell("0.15", "E2", xlContentType.Double);
            sh.AddCell(2, "A3", xlContentType.Integer);
            sh.AddCell(DateTime.Now, "B3", xlContentType.Date);
            sh.AddCell("Какая-то строка", "C3", xlContentType.SharedString);
            sh.AddCell("Какая-то строка", "D3", xlContentType.SharedString);
            sh.AddCell("0.25", "E3", xlContentType.Double);
            var memstream = new MemoryStream();
            xlWriter.Create(book).SaveToStream(memstream);

            qXlsx.FromStream(memstream, new qXlsxConfiguration { ValidationFailureEvent = (s, e) => { if (!e.MissingFields.Contains("Поле 1")) Assert.Fail(); } }).ReadToArray<TestExcelClass>();
            TestExcelClass[] data = qXlsx.FromStream(memstream).ReadToEnumerable<TestExcelClass>().ToArray();
        }
    }
}