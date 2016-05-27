using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLOC.Utility;
using XLOC.Utility.Events;
using DocumentFormat.OpenXml.Spreadsheet;
using XLOC.Utility.Extensions;
using System.ComponentModel;
using System.Reflection;

namespace XLOC.Reader
{
    class xlArrayReader
    {
        #region Var
        //=================================================
        XLOCConfiguration _config;
        DocDictionaries docProvider;
        //=================================================
        #endregion

        #region Events
        //=================================================
        public event EventHandler<SheetValidationErrorEventArgs> OnValidationFailure;
        public event EventHandler<CellReadingErrorEventArgs> OnCellReadingError;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        public xlArrayReader(XLOCConfiguration configuration)
        {
            _config = configuration;
            OnValidationFailure += _config.ValidationFailureEvent;
            OnValidationFailure += (s, e) => { };
            OnCellReadingError += _config.CellReadingErrorEvent;
            OnCellReadingError += (s, e) => { };
        }
        //=================================================
        #endregion

        #region Private
        //=================================================

        #region Sheet
        //=================================================
        protected object ConvertTypelessCell(Cell item)
        {
            if (item.StyleIndex != null && docProvider.styles[item.StyleIndex.Value].NumberFormatId.Value != 0 && item.CellValue != null)
            {
                switch (docProvider.styles[item.StyleIndex.Value].NumberFormatId.Value)
                {
                    case 1:
                    case 3:
                        return Convert.ToInt32(item.CellValue.Text);
                    case 2:
                    case 4:
                    case 11:
                    case 12:
                    case 49:
                    case 164:
                        return Convert.ToDecimal(item.CellValue.Text, new System.Globalization.CultureInfo("En"));
                    case 10:
                        return (Convert.ToDecimal(item.CellValue.Text, new System.Globalization.CultureInfo("En")) * 100).ToString("N2") + "%";
                    case 14:
                    case 165:
                    case 168:
                    case 169:
                    case 170:
                        return DateTime.FromOADate(double.Parse(item.CellValue.Text));
                    case 44:
                    case 167:
                        return (Convert.ToDecimal(item.CellValue.Text, new System.Globalization.CultureInfo("En"))).ToString("N2") + " ₽";
                    default:
                        throw new NotImplementedException($"Не реализован обработчик для формата {docProvider.styles[item.StyleIndex.Value].NumberFormatId.Value}");
                }
            }
            else
            {
                return !string.IsNullOrWhiteSpace(item.CellValue?.Text) ? (decimal?)Convert.ToDecimal(item.CellValue.Text, new System.Globalization.CultureInfo("En")) : null;
            }
        }
        //=================================================
        #endregion

        T RowToObject<T>(Row row, Map<T> map) where T : IxlCompatible, new()
        {
            T result = new T();
            foreach (var cell in map.GetCells(row))
            {
                try
                {
                    var mapItem = map[cell.CellReference.Value.rMatch(@"^[A-Z]+")];
                    mapItem.Property.SetValue(result, getValue(cell, mapItem.Property.PropertyType));
                }
                catch (Exception ex)
                {
                    OnCellReadingError(this, new CellReadingErrorEventArgs
                    {
                        Reference = cell.CellReference,
                        SourceType = cell.DataType?.Value ?? null,
                        OutputType = map[cell.CellReference.Value.rMatch("^[A-Z]+")].Attribute.ContentType,
                        Value = cell.CellValue?.Text,
                        Exception = ex
                    });
                    throw;
                }
            }
            return result;
        }

        object getValue(Cell cell, Type type)
        {
            try
            {
                switch (cell.DataType?.Value)
                {
                    case CellValues.Error:
                        throw new Exception(string.Format("Unknown cell type {0}", CellValues.Error));
                    case CellValues.Boolean:
                    case CellValues.Date:
                        throw new NotImplementedException($"Преобразование для типа {type} не реализовано");
                    case CellValues.Number:
                        return Convert.ChangeType(cell.CellValue, type);
                    case CellValues.SharedString:
                        var RefId = int.Parse(cell.CellValue.Text);
                        return TypeDescriptor.GetConverter(type).ConvertFromString(docProvider.sharedStrings[RefId].HasValue() ? docProvider.sharedStrings[RefId] : string.Empty);
                    case CellValues.String:
                    case CellValues.InlineString:
                        return Convert.ChangeType(cell.CellValue?.Text, type);
                    default:
                        return Convert.ChangeType(ConvertTypelessCell(cell), type);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Ошибка преобразования ячеек, адрес ссылки [{cell.CellReference}], искходное значение [{cell.CellValue?.Text}], исходный тип [{cell.DataType?.Value}], стиль [{cell.StyleIndex?.Value}]", ex);
            }
        }
        //=================================================
        #endregion

        #region Methods
        //=================================================
        public IEnumerable<T> ReadToEnumerable<T>(SpreadsheetDocument document) where T : IxlCompatible, new()
        {
            using (docProvider = new DocDictionaries(document))
            {
                var sheets = _config.Sheets == null ? document.WorkbookPart.Workbook.Sheets.Cast<Sheet>() : document.WorkbookPart.Workbook.Sheets.Cast<Sheet>().Where(x => _config.Sheets.Contains(x.SheetId.Value)).ToArray();
                foreach (var sheet in document.WorkbookPart.WorksheetParts.Where(x => sheets.Select(y => y.Id.Value).Contains(document.WorkbookPart.GetIdOfPart(x))))
                {
                    Map<T> map = GetMap<T>(sheet);

                    if (!map.IsValid)
                    {
                        OnValidationFailure(this, new SheetValidationErrorEventArgs(map.MissingFields, map.Exceptioins));
                        continue;
                    }

                    foreach (var row in sheet.Worksheet.Descendants<Row>().Skip(getSkip()))
                    {
                        T tmp;
                        try { tmp = RowToObject<T>(row, map); }
                        catch { if (_config.ContinueOnRowReadingError) continue; else throw; }
                        yield return tmp;
                    }
                }
            }
        }

        int getSkip() => _config.SkipMode == SkipModeEnum.None ? 0 : _config.SkipCount ?? 0;

        private Map<T> GetMap<T>(WorksheetPart sheet) where T : IxlCompatible, new()
        {
            switch (_config.SkipMode)
            {
                case SkipModeEnum.None: return new Map<T>(sheet.GetCaptionCells().ToDictionary(x => x.CellReference.Value, x => getValue(x, typeof(string)).ToString()));
                case SkipModeEnum.Manual: return new Map<T>(sheet.GetCaptionCells(_config.SkipCount.Value).ToDictionary(x => x.CellReference.Value, x => getValue(x, typeof(string)).ToString()));
                case SkipModeEnum.Auto: return AutoMap<T>(sheet);
                default: throw new NotImplementedException("Default switch case not implemented");
            }
        }

        private Map<T> AutoMap<T>(WorksheetPart sheet)
        {
            Map<T> result = null;
            var enumerator = sheet.Worksheet.GetFirstChild<SheetData>().Descendants<Row>().GetEnumerator();
            _config.SkipCount = 0;
            while (!(!enumerator.MoveNext() || (result?.IsValid ?? false)))
            {
                result = new Map<T>(ToDictionary(enumerator.Current));
                _config.SkipCount++;
            }
            return result;
        }

        private Dictionary<string, string> ToDictionary(Row row)
        {
            return row.Descendants<Cell>().Where(x => x.CellValue != null).ToDictionary(x => x.CellReference.Value, x => getValue(x, typeof(string)).ToString());
        }
        //=================================================
        #endregion
    }
}