using XLOC;
using XLOC.Utility;
using System;

namespace ExcelReaderUnitTestProject
{
    public class TestExcelClass
    {
        [XlField(XlContentType.Integer, "Поле 1")]
        public int intProperty1 { get; set; }
        [XlField(XlContentType.Integer, false, "Поле 2")]
        public int? intProperty2 { get; set; }
        [XlField(XlContentType.Integer, false, "Поле 3")]
        public int? intProperty3 { get; set; }
        [XlField(XlContentType.Double, "дробь")]
        public decimal decimalProperty { get; set; }
        [XlField(XlContentType.Date, "Какая-то дата")]
        public DateTime SomeDate { get; set; }
        [XlField(XlContentType.SharedString, false, "Какая-то строка")]
        public string SomeString { get; set; }
        [XlField(XlContentType.SharedString, false, "Мультизагаловок1", "Мультизагаловок2")]
        public string MultiCaption { get; set; }
        [XlField(XlContentType.SharedString, false, "GuidField")]
        public Guid Guid { get; set; }
    }
}