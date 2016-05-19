using XLOC;
using XLOC.Utility;
using System;

namespace ExcelReaderUnitTestProject
{
    public class TestExcelClass : IxlCompatible
    {
        [xlField(xlContentType.Integer, "Поле 1")]
        public int intProperty1 { get; set; }
        [xlField(xlContentType.Integer, false, "Поле 2")]
        public int? intProperty2 { get; set; }
        [xlField(xlContentType.Integer, false, "Поле 3")]
        public int? intProperty3 { get; set; }
        [xlField(xlContentType.Double, "дробь")]
        public decimal decimalProperty { get; set; }
        [xlField(xlContentType.Date, "Какая-то дата")]
        public DateTime SomeDate { get; set; }
        [xlField(xlContentType.SharedString, false, "Какая-то строка")]
        public string SomeString { get; set; }
        [xlField(xlContentType.SharedString, false, "Мультизагаловок1", "Мультизагаловок2")]
        public string MultiCaption { get; set; }
        [xlField(xlContentType.SharedString, false, "GuidField")]
        public Guid Guid { get; set; }
    }
}