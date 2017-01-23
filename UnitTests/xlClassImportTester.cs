using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using XLOC;
using XLOC.Book;
using XLOC.Utility;
using XLOC.Writer;

namespace ExcelReaderUnitTestProject
{
    [TestClass]
    public class xlClassImportTester
    {
        static string path = string.Format(@"{0}\{1}", Path.Combine(Environment.CurrentDirectory), "test2.xlsx");

        TestExcelClass[] data = new TestExcelClass[]
        {
            new TestExcelClass() { intProperty1 = 1 , intProperty3 = 1 , SomeDate = DateTime.Now, SomeString = "asdasd"},
            new TestExcelClass() { intProperty1 = 2 , intProperty3 = 2 , SomeDate = DateTime.Now, SomeString = "aafgf"},
            new TestExcelClass() { intProperty1 = 3 , intProperty3 = 3 , SomeDate = DateTime.Now, SomeString = "xdfe"},
            new TestExcelClass() { intProperty1 = 4 , intProperty3 = 4 , SomeDate = DateTime.Now, SomeString = "dfdr"},
            new TestExcelClass() { intProperty1 = 5 , intProperty3 = 5 , SomeDate = DateTime.Now, SomeString = "ghdg"},
            new TestExcelClass() { intProperty1 = 7 , intProperty3 = 7 , SomeDate = DateTime.Now, SomeString = "dfg"},
            new TestExcelClass() { intProperty1 = 9 , intProperty3 = 9 , SomeDate = DateTime.Now, SomeString = "dfgag"},
            new TestExcelClass() { intProperty1 = 10, intProperty3 = 10, SomeDate = DateTime.Now, SomeString = "sdfsw"},
        };

        [ClassInitialize]
        public static void Initialize(TestContext ctx) => File.Delete(path);

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
            var readedData = XlConverter.FromFile(path).ReadToEnumerable<TestExcelClass>().ToArray();
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
            using (var memstream = new MemoryStream())
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

                xlWriter.Create(book).SaveToStream(memstream);

                TestExcelClass[] data = XlConverter.FromStream(memstream).ReadToEnumerable<TestExcelClass>().ToArray();
                Assert.AreEqual(2, data.Count());
                Assert.IsTrue(data.All(x => !x.intProperty2.HasValue));
            }
        }

        [TestMethod]
        public void ReadToArrayWithNullableColumns()
        {
            using (var memstream = new MemoryStream())
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

                xlWriter.Create(book).SaveToStream(memstream);

                TestExcelClass[] data = XLOC.XlConverter.FromStream(memstream, new XLOCConfiguration { CellReadingErrorEvent = (s, e) => { throw new Exception(e.Exception.Message); } }).ReadToArray<TestExcelClass>();
                Assert.AreEqual(2, data.Count());
                Assert.IsTrue(data.All(x => !x.intProperty3.HasValue));
            }
        }

        [TestMethod]
        public void MultiCaptionTest()
        {
            int countShouldBe = 4;
            using (var memstream = new MemoryStream())
            {
                var book = new xlBook();
                var sheet = book.AddSheet("sheet1");
                sheet.AddCell("Поле 1", "A1", xlContentType.SharedString);
                sheet.AddCell("Какая-то дата", "B1", xlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", "C1", xlContentType.SharedString);
                sheet.AddCell("Мультизагаловок2", "D1", xlContentType.SharedString);
                sheet.AddCell("дробь", "E1", xlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", xlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", xlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", xlContentType.SharedString);
                    sheet.AddCell($"Какая-то строка{i}", $"D{i}", xlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", xlContentType.Double);
                }

                xlWriter.Create(book).SaveToStream(memstream);

                var isValid = true;
                TestExcelClass[] data = XlConverter.FromStream(memstream, new XLOCConfiguration { ValidationFailureEvent = (s, e) => isValid = false }).ReadToEnumerable<TestExcelClass>().ToArray();
                Assert.AreEqual(0, data.Count());
                Assert.IsFalse(isValid);
            }
        }

        [TestMethod]
        public void ValidationEventTest()
        {
            int countShouldBe = 4;
            using (var memstream = new MemoryStream())
            {
                var book = new xlBook();
                const string sheetName = "sheet1";
                var sheet = book.AddSheet(sheetName);

                sheet.AddCell("Какая-то дата", "B1", xlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", "C1", xlContentType.SharedString);
                sheet.AddCell("дробь", "E1", xlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell(DateTime.Now, $"B{i}", xlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", xlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", xlContentType.Double);
                }
                xlWriter.Create(book).SaveToStream(memstream);

                XlConverter.FromStream(memstream, new XLOCConfiguration { ValidationFailureEvent = (s, e) => { if (!e.MissingFields.Contains("Поле 1") || e.Sheet.Name != sheetName) Assert.Fail(); } }).ReadToArray<TestExcelClass>();
                TestExcelClass[] data = XlConverter.FromStream(memstream).ReadToEnumerable<TestExcelClass>().ToArray();
            }
        }

        [TestMethod]
        public void CellEventTest()
        {
            int countShouldBe = 4;
            using (var memstream = new MemoryStream())
            {
                var book = new xlBook();
                var sheet = book.AddSheet("sheet1");

                sheet.AddCell("Поле 1", "A1", xlContentType.SharedString);
                sheet.AddCell("Какая-то дата", "B1", xlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", "C1", xlContentType.SharedString);
                sheet.AddCell("дробь", "E1", xlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell("A", $"A{i}", xlContentType.SharedString);
                    sheet.AddCell(DateTime.Now, $"B{i}", xlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", xlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", xlContentType.Double);
                }
                xlWriter.Create(book).SaveToStream(memstream);
                bool result = true;
                XlConverter.FromStream(memstream, new XLOCConfiguration { CellReadingErrorEvent = (s, e) => { if (e.Reference != "A2") result = false; } }).ReadToArray<TestExcelClass>();
                Assert.IsFalse(result);
                TestExcelClass[] data = XlConverter.FromStream(memstream).ReadToEnumerable<TestExcelClass>().ToArray();
            }
        }

        [TestMethod]
        public void SkiperNone()
        {
            int countShouldBe = 4;
            using (var memstream = new MemoryStream())
            {
                var book = new xlBook();
                var sheet = book.AddSheet("test");

                sheet.AddCell("Поле 1", "A1", xlContentType.SharedString);
                sheet.AddCell("Какая-то дата", "B1", xlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", "C1", xlContentType.SharedString);
                sheet.AddCell("дробь", "E1", xlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", xlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", xlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", xlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", xlContentType.Double);
                }

                xlWriter.Create(book).SaveToStream(memstream);
                var conv = XlConverter.FromStream(memstream, new XLOCConfiguration { SkipMode = SkipModeEnum.None, SkipCount = 4 });
                Assert.AreEqual(countShouldBe, conv.ReadToArray<TestExcelClass>().Count());
            }
        }

        [TestMethod]
        public void SkiperManual()
        {
            int countShouldBe = 4;
            using (var memstream = new MemoryStream())
            {
                var book = new xlBook();
                var sheet = book.AddSheet("test");
                sheet.AddCell("Caption", "A1", xlContentType.SharedString);
                sheet.AddCell("Caption2", "A2", xlContentType.SharedString);

                sheet.AddCell("Поле 1", "A3", xlContentType.SharedString);
                sheet.AddCell("Какая-то дата", "B3", xlContentType.SharedString);
                sheet.AddCell("Мультизагаловок2", "D3", xlContentType.SharedString);
                sheet.AddCell("дробь", "E3", xlContentType.SharedString);

                for (int i = 4; i < 4 + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", xlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", xlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", xlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", xlContentType.Double);
                }

                xlWriter.Create(book).SaveToStream(memstream);
                var convertor = XlConverter.FromStream(memstream, new XLOCConfiguration { SkipMode = SkipModeEnum.Manual, SkipCount = 2 });
                Assert.AreEqual(countShouldBe, convertor.ReadToArray<TestExcelClass>().Count());
            }
        }

        [TestMethod]
        public void SkiperAuto()
        {
            int countShouldBe = 4;
            using (var memstream = new MemoryStream())
            {
                var book = new xlBook();
                var sheet = book.AddSheet("test");

                int skip = new Random().Next(4) + 1;
                for (int i = 0; i < skip; i++)
                {
                    sheet.AddCell($"Caption{i + 1}", $"A{i + 1}", xlContentType.SharedString);
                }

                sheet.AddCell("Поле 1", $"A{++skip}", xlContentType.SharedString);
                sheet.AddCell("Какая-то дата", $"B{skip}", xlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", $"C{skip}", xlContentType.SharedString);
                sheet.AddCell("дробь", $"E{skip}", xlContentType.SharedString);

                for (int i = ++skip; i < skip + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", xlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", xlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", xlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", xlContentType.Double);
                }

                xlWriter.Create(book).SaveToStream(memstream);
                var convertor = XlConverter.FromStream(memstream, new XLOCConfiguration { SkipMode = SkipModeEnum.Auto, SkipCount = 1 });
                Assert.AreEqual(countShouldBe, convertor.ReadToArray<TestExcelClass>().Count());
            }
        }

        [TestMethod]
        public void ExponentialNotice()
        {
            int countShouldBe = 4;
            using (var memstream = new MemoryStream())
            {
                var book = new xlBook();
                var sheet = book.AddSheet("test");

                sheet.AddCell("Поле 1", $"A1", xlContentType.SharedString);
                sheet.AddCell("Какая-то дата", $"B1", xlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", $"C1", xlContentType.SharedString);
                sheet.AddCell("дробь", $"AB1", xlContentType.SharedString);
                sheet.AddCell("noize", $"AC1", xlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", xlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", xlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", xlContentType.SharedString);
                    sheet.AddCell($"{(i / 100M).ToString("E")}", $"AB{i}", xlContentType.Double);
                    sheet.AddCell($"noize", $"AC{i}", xlContentType.Double);
                }

                xlWriter.Create(book).SaveToStream(memstream);
                var data = XlConverter.FromStream(memstream, new XLOCConfiguration { SkipMode = SkipModeEnum.Auto, ContinueOnRowReadingError = false }).ReadToArray<TestExcelClass>();
                Assert.AreEqual(countShouldBe, data.Count());
                Assert.IsTrue(data.All(x => x.decimalProperty != 0));
            }
        }

        [TestMethod]
        public void AutoDisposing()
        {
            int countShouldBe = 4;
            using (var ms = new MemoryStream())
            {
                var book = new xlBook();
                var sheet = book.AddSheet("test");

                sheet.AddCell("Поле 1", $"A1", xlContentType.SharedString);
                sheet.AddCell("Какая-то дата", $"B1", xlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", $"C1", xlContentType.SharedString);
                sheet.AddCell("дробь", $"AB1", xlContentType.SharedString);
                sheet.AddCell("noize", $"AC1", xlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", xlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", xlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", xlContentType.SharedString);
                    sheet.AddCell($"{(i / 100M).ToString("E")}", $"AB{i}", xlContentType.Double);
                    sheet.AddCell($"noize", $"AC{i}", xlContentType.Double);
                }

                xlWriter.Create(book).SaveToStream(ms);
                var data = XlConverter.FromStream(ms, new XLOCConfiguration { SkipMode = SkipModeEnum.Auto, ContinueOnRowReadingError = false, AutoDispose = false }).ReadToEnumerable<TestExcelClass>();
                Assert.AreEqual(countShouldBe, data.Count());
                Assert.IsTrue(data.All(x => x.decimalProperty != 0));
            }
        }
    }
}