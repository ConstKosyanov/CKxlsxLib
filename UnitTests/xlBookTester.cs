using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using System.Linq;
using XLOC.Book;
using XLOC.Utility;
using XLOC.Writer;

namespace ExcelReaderUnitTestProject
{
    [TestClass]
    public class xlBookTester
    {
        string list1 = "SharedStringTable";
        string list2 = "FormatedCells";
        xlBook xl = new xlBook();

        [ClassInitialize]
        public static void Initialize(TestContext ctx)
        {
            File.Delete(string.Format(@"{0}\{1}", Path.Combine(Environment.CurrentDirectory), "test.xlsx"));
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void NameValidation()
        {
            xl.Name = "test:";
        }

        [TestMethod]
        public void NameSet()
        {
            xl.Name = "test.asd1 ";
            Assert.AreEqual("test.xlsx", xl.Name, "Не верно обратонано расширение");
        }

        [TestMethod]
        public void AddSheet()
        {
            Assert.IsNotNull(xl.Sheets, "Список листов не определён");
            xl.AddSheet(list2);
            Assert.AreEqual(1, xl.Sheets.Count());
            xl[0].AddCell(1, "A1",xlContentType.Integer);
            xl[0].AddCell(DateTime.Now, "A2", xlContentType.Date);
            xl[0].AddCell(3.14, "A3", xlContentType.Double);

            xl.AddSheet(list1);
            var rnd = new Random();
            xl[1].AddCell("Test", 1, 1, xlContentType.SharedString);
            xl[1].AddCell("Test2", 2, 1, xlContentType.SharedString);
            xl[1].AddCell("Test", 1, 2, xlContentType.SharedString);
            xl[1].AddCell("Test3", 2, 2, xlContentType.SharedString);

            Assert.AreEqual(2, xl.Sheets.Count());
            Assert.AreEqual(list1, xl[1].Name, "Не задано имя листа");
        }

        [TestMethod]
        public void AddSheetAndSave()
        {
            NameSet();
            AddSheet();
            SaveBookThrougStream();
            ReadBook(xl.Name);
            Assert.AreEqual(2, xl.Sheets.Count(), "Количество листов не совпадает");
            Assert.AreEqual(list1, xl[1].Name, "Имя листа прочитано не верно");
            Assert.AreEqual("Test", xl[1].Get(1, 1).Value.ToString(), "Чтение из SharedStringTable реализовано не верно");
            Assert.AreEqual("Test2", xl[1].Get(1, 2).Value.ToString(), "Чтение из SharedStringTable реализовано не верно");
            Assert.AreEqual("Test", xl[1].Get(2, 1).Value.ToString(), "Чтение из SharedStringTable реализовано не верно");
            Assert.AreEqual("Test3", xl[1].Get(2, 2).Value.ToString(), "Чтение из SharedStringTable реализовано не верно");

            Assert.AreEqual(Convert.ToDecimal(1), xl[0].Get(1, 1).Value, "Чтение даты реализовано не верно");
            Assert.AreEqual(DateTime.Today, ((DateTime)xl[0].Get(1, 2).Value).Date, "Чтение даты реализовано не верно");
            Assert.AreEqual((decimal)3.14, xl[0].Get(1, 3).Value, "Чтение даты реализовано не верно");
        }

        [TestMethod]
        public void SaveBookThrougStream()
        {
            if (string.IsNullOrWhiteSpace(xl.Name))
                NameSet();
            if (xl.Sheets.Count() == 0)
                AddSheet();
            try
            {
                using (var file = File.Create(string.Format(@"{0}\{1}", Path.Combine(Environment.CurrentDirectory), xl.Name)))
                {
                    var writer = xlWriter.Create(xl);
                    var err = writer.SaveToStream(file);
                    if (err.Count() > 0)
                        Assert.Fail("Ошибка сохранения:\n{0}", string.Join("\n", err.Select(x => x.Description)));
                }
            }
            catch (Exception ex) { Assert.Fail(string.Format("Ошибка сохранения\n{0}", ex.Message)); }
        }

        [TestMethod]
        public void SaveBookThroughFile()
        {
            if (string.IsNullOrWhiteSpace(xl.Name))
                NameSet();
            if (xl.Sheets.Count() == 0)
                AddSheet();
            try
            {
                var err = xlWriter.Create(xl).SaveToFile(string.Format(@"{0}\{1}", Path.Combine(Environment.CurrentDirectory), xl.Name));
                if (err.Count() > 0)
                    Assert.Fail("Ошибка сохранения:\n{0}", string.Join("\n", err.Select(x => x.Description)));
            }
            catch (Exception ex) { Assert.Fail(string.Format("Ошибка сохранения\n{0}", ex.Message)); }
        }

        public void ReadBook(string path)
        {
            try
            {
                using (var file = File.Open(path, FileMode.Open))
                {
                    var streamReader = XLOC.XlConverter.FromStream(file);
                    xl = streamReader.ReadToBook();
                }

                var fileReader = XLOC.XlConverter.FromFile(path);
                xl = fileReader.ReadToBook();
            }
            catch (Exception ex) { Assert.Fail(string.Format("Ошибка чтения\n{0}", ex.Message)); }
        }

        [ClassCleanup]
        public static void Cleanup()
        {
            //System.Diagnostics.Process.Start(string.Format(@"{0}\{1}", Path.Combine(Environment.CurrentDirectory), "test.xlsx"));
        }
    }
}