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
    internal class xlBookReader : xlReader
    {
        #region Constructor
        //=================================================
        internal xlBookReader(XLOCConfiguration configuration) : base(configuration) { }
        //=================================================
        #endregion

        #region Methods
        //=================================================
        static xlContentType ToxlContentType(CellValues? val)
        {
            switch (val)
            {
                case CellValues.Boolean: return xlContentType.Boolean;
                case CellValues.Date: return xlContentType.Date;
                case CellValues.Error: throw new Exception(string.Format("Unknown cell type {0}", CellValues.Error));
                case CellValues.InlineString: throw new Exception(string.Format("Unknown cell type {0}", CellValues.InlineString));
                case CellValues.Number: return xlContentType.Double;
                case CellValues.SharedString: return xlContentType.SharedString;
                case CellValues.String: return xlContentType.String;
                default: return xlContentType.Void;
            }
        }

        protected xlCell ReadCell(Cell item)
        {
            string Reference = string.Empty;
            object Value = null;
            int? RefId = null;

            Reference = item.CellReference.Value;
            xlContentType? Type = ToxlContentType(item.DataType != null ? (CellValues?)item.DataType.Value : null);
            try
            {
                switch (Type)
                {
                    case xlContentType.Boolean:
                    case xlContentType.Date:
                        throw new NotImplementedException($"Преобразование для типа {Type} не реализовано");
                    case xlContentType.Double:
                        Value = item.CellValue != null
                            ? (decimal?)Convert.ToDecimal(item.CellValue.Text, new System.Globalization.CultureInfo("En"))
                            : null;
                        break;
                    case xlContentType.SharedString:
                        RefId = int.Parse(item.CellValue.Text);
                        Value = _config.DocProvider.sharedStrings[RefId.Value].HasValue()
                            ? _config.DocProvider.sharedStrings[RefId.Value]
                            : string.Empty;
                        break;
                    case xlContentType.String:
                        Value = (item.CellValue == null ? null : item.CellValue.Text);
                        break;
                    default:
                        Value = ConvertTypelessCell(item);
                        break;
                }
                return new xlCell(Reference, Type, Value, RefId);
            }
            catch (Exception ex)
            {
                cellErrorEventCaller(new CellReadingErrorEventArgs
                {
                    Reference = item.CellReference,
                    SourceType = item.DataType?.Value ?? null,
                    OutputType = xlContentType.Void,
                    Value = item.CellValue?.Text,
                    Sheet = null,
                    Exception = ex
                });
                throw new FormatException(string.Format("Ошибка конвертирования ячейки при чтении\nисходное значение: {0}\nконечный тип: {1}", item.CellValue.Text, Type), ex);
            }
        }

        public xlBook ReadToBook(SpreadsheetDocument document)
        {
            xlBook result = new xlBook();
            for (int i = 0; i < document.WorkbookPart.WorksheetParts.Count(); i++)
            {
                result.AddSheet(_config.DocProvider.sheetNames[i]);
                var sheet = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().SingleOrDefault(s => s.Name == _config.DocProvider.sheetNames[i]);
                var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
                foreach (var cell in worksheetPart.Worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
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