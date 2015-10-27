using System;

namespace CKxlsxLib
{
    public enum xlContentType : byte
    {
        Void,
        Boolean,
        Integer,
        Double,
        SharedString,
        String,
        Date,
    }

    public static class xlContentTypeMethods
    {
        public static xlContentType ToxlContentType(this DocumentFormat.OpenXml.Spreadsheet.CellValues? local)
        {
            switch (local)
            {
                case DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean:
                    return xlContentType.Boolean;
                case DocumentFormat.OpenXml.Spreadsheet.CellValues.Date:
                    return xlContentType.Date;
                case DocumentFormat.OpenXml.Spreadsheet.CellValues.Error:
                    throw new Exception(string.Format("Unknown cell type {0}", DocumentFormat.OpenXml.Spreadsheet.CellValues.Error));
                case DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString:
                    throw new Exception(string.Format("Unknown cell type {0}", DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString));
                case DocumentFormat.OpenXml.Spreadsheet.CellValues.Number:
                    return xlContentType.Double;
                case DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString:
                    return xlContentType.SharedString;
                case DocumentFormat.OpenXml.Spreadsheet.CellValues.String:
                    return xlContentType.String;
                default:
                    return xlContentType.Void;
            }
        }
    }
}