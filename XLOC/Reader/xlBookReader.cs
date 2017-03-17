using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLOC.Book;
using System.Linq;
using System;
using XLOC.Utility;
using XLOC.Utility.Extensions;
using XLOC.Utility.Events;

namespace XLOC.Reader
{
    internal class XlBookReader : XlReader
    {
        #region Constructor
        //=================================================
        internal XlBookReader(XLOCConfiguration configuration) : base(configuration) { }
        //=================================================
        #endregion

        #region Methods
        //=================================================
        static XlContentType toxlContentType(CellValues? val)
        {
            switch (val)
            {
                case CellValues.Boolean: return XlContentType.Boolean;
                case CellValues.Date: return XlContentType.Date;
                case CellValues.Error: throw new Exception(string.Format("Unknown cell type {0}", CellValues.Error));
                case CellValues.InlineString: throw new Exception(string.Format("Unknown cell type {0}", CellValues.InlineString));
                case CellValues.Number: return XlContentType.Double;
                case CellValues.SharedString: return XlContentType.SharedString;
                case CellValues.String: return XlContentType.String;
                default: return XlContentType.Void;
            }
        }

        protected XlCell ReadCell(Cell item)
        {
            string Reference = string.Empty;
            object Value = null;
            int? RefId = null;

            Reference = item.CellReference.Value;
            XlContentType? Type = toxlContentType(item.DataType != null ? (CellValues?)item.DataType.Value : null);
            try
            {
                switch (Type)
                {
                    case XlContentType.Boolean:
                    case XlContentType.Date:
                        throw new NotImplementedException($"Преобразование для типа {Type} не реализовано");
                    case XlContentType.Double:
                        Value = item.CellValue != null
                            ? (decimal?)Convert.ToDecimal(item.CellValue.Text, new System.Globalization.CultureInfo("En"))
                            : null;
                        break;
                    case XlContentType.SharedString:
                        RefId = int.Parse(item.CellValue.Text);
                        Value = _config.DocProvider.sharedStrings[RefId.Value].HasValue()
                            ? _config.DocProvider.sharedStrings[RefId.Value]
                            : string.Empty;
                        break;
                    case XlContentType.String:
                        Value = item.CellValue?.Text;
                        break;
                    default:
                        Value = ConvertTypelessCell(item);
                        break;
                }
                return new XlCell(Reference, Type, Value, RefId);
            }
            catch (Exception ex)
            {
                cellErrorEventCaller(new CellReadingErrorEventArgs
                {
                    Reference = item.CellReference,
                    SourceType = item.DataType?.Value ?? null,
                    OutputType = XlContentType.Void,
                    Value = item.CellValue?.Text,
                    Sheet = null,
                    Exception = ex
                });
                throw new FormatException(string.Format("Ошибка конвертирования ячейки при чтении\nисходное значение: {0}\nконечный тип: {1}", item.CellValue.Text, Type), ex);
            }
        }

        public XlBook ReadToBook(SpreadsheetDocument document)
        {
            XlBook result = new XlBook();
            for (int i = 0; i < document.WorkbookPart.WorksheetParts.Count(); i++)
            {
                result.AddSheet(_config.DocProvider.sheetNames[i]);
                Sheet sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().SingleOrDefault(s => s.Name == _config.DocProvider.sheetNames[i]);
                var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
                foreach (Cell cell in worksheetPart.Worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
                {
                    try
                    {
                        if (cell == null) { continue; }
                        result[i].Cells.Add(ReadCell(cell));
                    }
                    catch (Exception)
                    {
                        if (_config.ContinueOnRowReadingError) continue; else throw;
                    }
                }
            }
            return result;

        }
        //=================================================
        #endregion
    }
}