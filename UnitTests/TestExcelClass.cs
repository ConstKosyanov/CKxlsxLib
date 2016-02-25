using CKxlsxLib;
using System;

namespace ExcelReaderUnitTestProject
{
    public class TestExcelClass : IxlCompatible
    {
        [xlField(xlContentType.Integer, "Поле 1")]
        public int intProperty1 { get; set; }
        [xlField(xlContentType.Integer, "Поле 2")]
        public int? intProperty2 { get; set; }
        [xlField(xlContentType.Double,"дробь")]
        public decimal decimalProperty { get; set; }
        [xlField(xlContentType.Date, "Какая-то дата")]
        public DateTime SomeDate { get; set; }
        [xlField(xlContentType.SharedString, false, "Какая-то строка")]
        public string SomeString { get; set; }
        [xlField(xlContentType.SharedString, "Мультизагаловок1", "Мультизагаловок2")]
        public string MultiCaption { get; set; }
    }
}