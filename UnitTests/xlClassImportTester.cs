using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLOC;
using XLOC.Book;
using XLOC.Utility;
using XLOC.Writer;

namespace ExcelReaderUnitTestProject
{
    [TestClass]
    public class XlClassImportTester
    {
        #region Init
        //=================================================
        static string path = string.Format(@"{0}\{1}", Path.Combine(Environment.CurrentDirectory), "test2.xlsx");
        Random rnd = new Random();

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

        List<FieldGenerator> fields = new List<FieldGenerator>
        {
            new FieldGenerator { Value = FieldsEnum.intProperty2, contentType = XlContentType.Integer, Caption = "Поле 2", Filler = (x)=>x },
            new FieldGenerator { Value = FieldsEnum.intProperty3, contentType= XlContentType.Integer, Caption = "Поле 3", Filler = (x)=>x },
            new FieldGenerator { Value = FieldsEnum.decimalProperty, contentType= XlContentType.Double, Caption = "дробь", Filler = (x)=>$"0.{x}" },
            new FieldGenerator { Value = FieldsEnum.SomeDate, contentType= XlContentType.Date , Caption = "Какая-то дата", Filler = (x)=>DateTime.Now },
            new FieldGenerator { Value = FieldsEnum.SomeString, contentType = XlContentType.SharedString, Caption = "Какая-то строка", Filler = (x)=>$"Какая-то строка{x}" },
            new FieldGenerator { Value = FieldsEnum.Guid, contentType = XlContentType.SharedString, Caption = "GuidField", Filler = (x)=>Guid.NewGuid() },
        };

        [ClassInitialize]
        public static void Initialize(TestContext ctx) => File.Delete(path);
        //=================================================
        #endregion

        [TestMethod]
        public void Write()
        {
            DocumentFormat.OpenXml.Validation.ValidationErrorInfo[] err = XlConverter.FromEnumerable(data).SaveToFile(path);
            if (err.Count() > 0)
                Assert.Fail("Ошибка сохранения:\n{0}", string.Join("\n", err.Select(x => x.Description)));
        }

        [TestMethod]
        public void Read()
        {
            Write();
            TestExcelClass[] readedData = XlConverter.FromFile(path).ReadToEnumerable<TestExcelClass>().ToArray();
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
                var book = new XlBook();
                XlSheet sh = book.AddSheet("sheet1");

                #region Captions
                //=================================================
                sh.AddCell("Поле 1", "A1", XlContentType.SharedString);
                sh.AddCell("Какая-то дата", "B1", XlContentType.SharedString);
                sh.AddCell("Мультизагаловок2", "C1", XlContentType.SharedString);
                sh.AddCell("дробь", "E1", XlContentType.SharedString);
                //=================================================
                #endregion

                #region Data
                //=================================================
                sh.AddCell(1, "A2", XlContentType.Integer);
                sh.AddCell(DateTime.Now, "B2", XlContentType.Date);
                sh.AddCell("Какая-то строка", "C2", XlContentType.SharedString);
                sh.AddCell("0.15", "E2", XlContentType.Double);

                sh.AddCell(2, "A3", XlContentType.Integer);
                sh.AddCell(DateTime.Now, "B3", XlContentType.Date);
                sh.AddCell("Какая-то строка", "C3", XlContentType.SharedString);
                sh.AddCell("0.25", "E3", XlContentType.Double);
                //=================================================
                #endregion

                XlConverter.FromBook(book).SaveToStream(memstream);

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
                var book = new XlBook();
                XlSheet sh = book.AddSheet("sheet1");

                #region Captions
                //=================================================
                sh.AddCell("Поле 1", "A1", XlContentType.SharedString);
                sh.AddCell("Какая-то дата", "B1", XlContentType.SharedString);
                sh.AddCell("Мультизагаловок2", "C1", XlContentType.SharedString);
                sh.AddCell("дробь", "E1", XlContentType.SharedString);
                sh.AddCell("Поле 3", "F1", XlContentType.SharedString);
                //=================================================
                #endregion

                #region Data
                //=================================================
                sh.AddCell(1, "A2", XlContentType.SharedString);
                sh.AddCell(DateTime.Now, "B2", XlContentType.Date);
                sh.AddCell("Какая-то строка", "C2", XlContentType.SharedString);
                sh.AddCell("0.15", "E2", XlContentType.Double);
                sh.AddCell("", "F2", XlContentType.SharedString);

                sh.AddCell(2, "A3", XlContentType.Integer);
                sh.AddCell(DateTime.Now, "B3", XlContentType.Date);
                sh.AddCell("Какая-то строка", "C3", XlContentType.SharedString);
                sh.AddCell("0.25", "E3", XlContentType.Double);
                sh.AddCell("", "F3", XlContentType.SharedString);
                //=================================================
                #endregion

                XlConverter.FromBook(book).SaveToStream(memstream);

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
                var book = new XlBook();
                XlSheet sheet = book.AddSheet("sheet1");
                sheet.AddCell("Поле 1", "A1", XlContentType.SharedString);
                sheet.AddCell("Какая-то дата", "B1", XlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", "C1", XlContentType.SharedString);
                sheet.AddCell("Мультизагаловок2", "D1", XlContentType.SharedString);
                sheet.AddCell("дробь", "E1", XlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", XlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", XlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", XlContentType.SharedString);
                    sheet.AddCell($"Какая-то строка{i}", $"D{i}", XlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", XlContentType.Double);
                }

                XlConverter.FromBook(book).SaveToStream(memstream);

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
                var book = new XlBook();
                const string sheetName = "sheet1";
                XlSheet sheet = book.AddSheet(sheetName);

                sheet.AddCell("Какая-то дата", "B1", XlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", "C1", XlContentType.SharedString);
                sheet.AddCell("дробь", "E1", XlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell(DateTime.Now, $"B{i}", XlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", XlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", XlContentType.Double);
                }
                XlConverter.FromBook(book).SaveToStream(memstream);

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
                var book = new XlBook();
                XlSheet sheet = book.AddSheet("sheet1");

                sheet.AddCell("Поле 1", "A1", XlContentType.SharedString);
                sheet.AddCell("Какая-то дата", "B1", XlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", "C1", XlContentType.SharedString);
                sheet.AddCell("дробь", "E1", XlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell("A", $"A{i}", XlContentType.SharedString);
                    sheet.AddCell(DateTime.Now, $"B{i}", XlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", XlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", XlContentType.Double);
                }
                XlConverter.FromBook(book).SaveToStream(memstream);
                bool result = true;
                XlConverter.FromStream(memstream, new XLOCConfiguration { CellReadingErrorEvent = (s, e) => { if (e.Reference != "A2") result = false; }, AutoDispose = false }).ReadToArray<TestExcelClass>();
                Assert.IsFalse(result);
                TestExcelClass[] data = XlConverter.FromStream(memstream, new XLOCConfiguration { AutoDispose = true }).ReadToArray<TestExcelClass>();
            }
        }

        [TestMethod]
        public void SkiperNone()
        {
            int countShouldBe = 4;
            using (var memstream = new MemoryStream())
            {
                var book = new XlBook();
                XlSheet sheet = book.AddSheet("test");

                sheet.AddCell("Поле 1", "A1", XlContentType.SharedString);
                sheet.AddCell("Какая-то дата", "B1", XlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", "C1", XlContentType.SharedString);
                sheet.AddCell("дробь", "E1", XlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", XlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", XlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", XlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", XlContentType.Double);
                }

                XlConverter.FromBook(book).SaveToStream(memstream);
                XLOCReader conv = XlConverter.FromStream(memstream, new XLOCConfiguration { SkipMode = SkipModeEnum.None, SkipCount = 4 });
                Assert.AreEqual(countShouldBe, conv.ReadToArray<TestExcelClass>().Count());
            }
        }

        [TestMethod]
        public void SkiperManual()
        {
            int countShouldBe = 4;
            using (var memstream = new MemoryStream())
            {
                var book = new XlBook();
                XlSheet sheet = book.AddSheet("test");
                sheet.AddCell("Caption", "A1", XlContentType.SharedString);
                sheet.AddCell("Caption2", "A2", XlContentType.SharedString);

                sheet.AddCell("Поле 1", "A3", XlContentType.SharedString);
                sheet.AddCell("Какая-то дата", "B3", XlContentType.SharedString);
                sheet.AddCell("Мультизагаловок2", "D3", XlContentType.SharedString);
                sheet.AddCell("дробь", "E3", XlContentType.SharedString);

                for (int i = 4; i < 4 + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", XlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", XlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", XlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", XlContentType.Double);
                }

                XlConverter.FromBook(book).SaveToStream(memstream);
                XLOCReader convertor = XlConverter.FromStream(memstream, new XLOCConfiguration { SkipMode = SkipModeEnum.Manual, SkipCount = 2 });
                Assert.AreEqual(countShouldBe, convertor.ReadToArray<TestExcelClass>().Count());
            }
        }

        [TestMethod]
        public void SkiperAuto()
        {
            int countShouldBe = 4;
            using (var memstream = new MemoryStream())
            {
                var book = new XlBook();
                XlSheet sheet = book.AddSheet("test");

                int skip = rnd.Next(4) + 1;
                for (int i = 0; i < skip; i++)
                {
                    sheet.AddCell($"Caption{i + 1}", $"A{i + 1}", XlContentType.SharedString);
                }

                sheet.AddCell("Поле 1", $"A{++skip}", XlContentType.SharedString);
                sheet.AddCell("Какая-то дата", $"B{skip}", XlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", $"C{skip}", XlContentType.SharedString);
                sheet.AddCell("дробь", $"E{skip}", XlContentType.SharedString);

                for (int i = ++skip; i < skip + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", XlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", XlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", XlContentType.SharedString);
                    sheet.AddCell($"0.1{i}", $"E{i}", XlContentType.Double);
                }

                XlConverter.FromBook(book).SaveToStream(memstream);
                XLOCReader convertor = XlConverter.FromStream(memstream, new XLOCConfiguration { SkipMode = SkipModeEnum.Auto, SkipCount = 1 });
                Assert.AreEqual(countShouldBe, convertor.ReadToArray<TestExcelClass>().Count());
            }
        }

        [TestMethod]
        public void ExponentialNotice()
        {
            int countShouldBe = 4;
            using (var memstream = new MemoryStream())
            {
                var book = new XlBook();
                XlSheet sheet = book.AddSheet("test");

                sheet.AddCell("Поле 1", $"A1", XlContentType.SharedString);
                sheet.AddCell("Какая-то дата", $"B1", XlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", $"C1", XlContentType.SharedString);
                sheet.AddCell("дробь", $"AB1", XlContentType.SharedString);
                sheet.AddCell("noize", $"AC1", XlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", XlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", XlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", XlContentType.SharedString);
                    sheet.AddCell($"{(i / 100M).ToString("E")}", $"AB{i}", XlContentType.Double);
                    sheet.AddCell($"noize", $"AC{i}", XlContentType.Double);
                }

                XlConverter.FromBook(book).SaveToStream(memstream);
                TestExcelClass[] data = XlConverter.FromStream(memstream, new XLOCConfiguration { SkipMode = SkipModeEnum.Auto, ContinueOnRowReadingError = false }).ReadToArray<TestExcelClass>();
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
                var book = new XlBook();
                XlSheet sheet = book.AddSheet("test");

                sheet.AddCell("Поле 1", $"A1", XlContentType.SharedString);
                sheet.AddCell("Какая-то дата", $"B1", XlContentType.SharedString);
                sheet.AddCell("Мультизагаловок1", $"C1", XlContentType.SharedString);
                sheet.AddCell("дробь", $"AB1", XlContentType.SharedString);
                sheet.AddCell("noize", $"AC1", XlContentType.SharedString);

                for (int i = 2; i < 2 + countShouldBe; i++)
                {
                    sheet.AddCell(i, $"A{i}", XlContentType.Integer);
                    sheet.AddCell(DateTime.Now, $"B{i}", XlContentType.Date);
                    sheet.AddCell($"Какая-то строка{i}", $"C{i}", XlContentType.SharedString);
                    sheet.AddCell($"{(i / 100M).ToString("E")}", $"AB{i}", XlContentType.Double);
                    sheet.AddCell($"noize", $"AC{i}", XlContentType.Double);
                }

                XlConverter.FromBook(book).SaveToStream(ms);
                IEnumerable<TestExcelClass> data = XlConverter.FromStream(ms, new XLOCConfiguration { SkipMode = SkipModeEnum.Auto, ContinueOnRowReadingError = false, AutoDispose = false }).ReadToEnumerable<TestExcelClass>();
                Assert.AreEqual(countShouldBe, data.Count());
                Assert.IsTrue(data.All(x => x.decimalProperty != 0));
            }
        }

        #region Group
        //=================================================
        [TestMethod]
        public void ReadToGroup_SameColumns()
        {
            Dictionary<int, int> result = new Dictionary<int, int> { };

            using (var ms = new MemoryStream())
            {
                var book = new XlBook();

                for (int i = 0; i < rnd.Next(3, 5); i++)
                {
                    XlSheet sheet = book.AddSheet($"Sheet{i}");

                    sheet.AddCell("Поле 1", "A1", XlContentType.SharedString);
                    sheet.AddCell("Какая-то дата", "B1", XlContentType.SharedString);
                    sheet.AddCell("Мультизагаловок1", "C1", XlContentType.SharedString);
                    sheet.AddCell("дробь", "E1", XlContentType.SharedString);

                    result[i] = rnd.Next(4, 6);
                    for (int j = 2; j < 2 + result[i]; j++)
                    {
                        sheet.AddCell(j, $"A{j}", XlContentType.Integer);
                        sheet.AddCell(DateTime.Now, $"B{j}", XlContentType.Date);
                        sheet.AddCell($"Какая-то строка{j}", $"C{j}", XlContentType.SharedString);
                        sheet.AddCell($"0.1{j}", $"E{j}", XlContentType.Double);
                    }
                }

                XlConverter.FromBook(book).SaveToStream(ms);
                IEnumerable<IGrouping<SheetIdentifier, TestExcelClass>> data = XlConverter.FromStream(ms, new XLOCConfiguration { }).ReadToGroup<TestExcelClass>();

                Assert.AreEqual(result.Count, data.Count(), "Количество листов не совпадает");
                foreach (KeyValuePair<int, int> item in result)
                {
                    IGrouping<SheetIdentifier, TestExcelClass> itemsCount = data.Single(x => x.Key.Name == $"Sheet{item.Key}");
                    Assert.AreEqual(item.Value, itemsCount.Count(), $"Количество записей на листе {item.Key} не соответсвует");
                }
            }
        }

        [TestMethod]
        public void ReadToGroup_SameClass()
        {
            Dictionary<int, int> result = new Dictionary<int, int> { };

            using (var ms = new MemoryStream())
            {
                var book = new XlBook();

                for (int i = 0; i < rnd.Next(3, 5); i++)
                {
                    XlSheet sheet = book.AddSheet($"Sheet{i}");

                    FieldGenerator[] columns = fields.Where(x => rnd.Next(0, 5) < 5).ToArray();
                    var letter = 'B';
                    sheet.AddCell("Поле 1", "A1", XlContentType.SharedString);
                    foreach (FieldGenerator item in columns)
                        sheet.AddCell(item.Caption, $"{item.Col = letter++}1", XlContentType.SharedString);


                    result[i] = rnd.Next(4, 6);
                    for (int j = 2; j < 2 + result[i]; j++)
                        foreach (FieldGenerator item in columns)
                            sheet.AddCell(item.Filler(j), $"{item.Col}{j}", item.contentType);
                }

                XlConverter.FromBook(book).SaveToStream(ms);
                IEnumerable<IGrouping<SheetIdentifier, TestExcelClass>> data = XlConverter.FromStream(ms, new XLOCConfiguration { }).ReadToGroup<TestExcelClass>();

                Assert.AreEqual(result.Count, data.Count(), "Количество листов не совпадает");
                foreach (KeyValuePair<int, int> item in result)
                {
                    IGrouping<SheetIdentifier, TestExcelClass> itemsCount = data.Single(x => x.Key.Name == $"Sheet{item.Key}");
                    Assert.AreEqual(item.Value, itemsCount.Count(), $"Количество записей на листе {item.Key} не соответсвует");
                }
            }
        }

        [TestMethod]
        public void ReadToGroup_DifferentClasses()
        {
            using (var ms = new MemoryStream())
            {
                var book = new XlBook();

                XlSheet sheet = book.AddSheet("Sheet1");

                char col = 'B';
                sheet.AddCell("Поле 1", "A1", XlContentType.SharedString);
                foreach (FieldGenerator item in fields)
                    sheet.AddCell(item.Caption, $"{item.Col = col++}1", XlContentType.SharedString);

                for (int i = 2; i < 5; i++)
                    foreach (FieldGenerator item in fields)
                        sheet.AddCell(item.Filler(i), $"{item.Col}{i}", item.contentType);

                sheet = book.AddSheet("Sheet2");
                sheet.AddCell("MyProperty", "A1", XlContentType.SharedString);
                sheet.AddCell("Test2UniqueField", "B1", XlContentType.SharedString);

                for (int i = 2; i < 5; i++)
                {
                    sheet.AddCell(i, $"A{i}", XlContentType.Integer);
                    sheet.AddCell(i * 2, $"B{i}", XlContentType.Integer);
                }

                XlConverter.FromBook(book).SaveToStream(ms);
                XLOCReader reader = XlConverter.FromStream(ms, new XLOCConfiguration { });

                IEnumerable<IGrouping<SheetIdentifier, TestExcelClass>> res1 = reader.ReadToGroup<TestExcelClass>();
                Assert.AreEqual(1, res1.Count(x => x.Any()));
                Assert.AreEqual(3, res1.Single(x => x.Any()).Count());

                IEnumerable<IGrouping<SheetIdentifier, TestExcelClass2>> res2 = reader.ReadToGroup<TestExcelClass2>();
                Assert.AreEqual(1, res2.Count(x => x.Any()));
                Assert.AreEqual(3, res2.Single(x => x.Any()).Count());

                Assert.AreEqual(3, reader.ReadToEnumerable<TestExcelClass>().Count());
                Assert.AreEqual(3, reader.ReadToEnumerable<TestExcelClass2>().Count());
            }
        }
        //=================================================
        #endregion

        #region Sub classes
        //=================================================
        enum FieldsEnum { intProperty2, intProperty3, decimalProperty, SomeDate, SomeString, Guid }

        class FieldGenerator
        {
            public FieldsEnum Value { get; set; }
            public XlContentType contentType { get; set; }
            public Func<int, object> Filler { get; set; }
            public string Caption { get; set; }
            public char Col { get; set; }
        }

        class TestExcelClass2
        {
            [XlField]
            public int MyProperty { get; set; }
            [XlField]
            public int Test2UniqueField { get; set; }
        }
        //=================================================
        #endregion
    }
}