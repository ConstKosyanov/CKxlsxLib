using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.ComponentModel;
using System.IO;
using System.Linq;
using XLOC.Book;
using XLOC.Utility;
using XLOC.Utility.Events;
using XLOC.Utility.Extensions;

namespace XLOC.Reader
{
    public abstract class xlReader
    {
        #region Variables
        //=================================================
        protected XLOCConfiguration _config;
        protected DocDictionaries _docProvider;
        //=================================================
        #endregion

        #region Events
        //=================================================
        public event EventHandler<CellReadingErrorEventArgs> OnCellReadingError;

        protected void cellErrorEventCaller(CellReadingErrorEventArgs args) => OnCellReadingError(this, args);
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        protected xlReader(XLOCConfiguration configuration)
        {
            _config = configuration;
            OnCellReadingError += (s, e) => { };
            OnCellReadingError += _config.CellReadingErrorEvent;
        }
        //=================================================
        #endregion

        #region private
        //=================================================
        static object ConvertToTypeWitNullableCheck(object value, Type type) => type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>) ? ConvertNullable(value, type) : Convert.ChangeType(value, type);

        static object ConvertNullable(object value, Type type) => value != null ? Convert.ChangeType(value, type.GetGenericArguments().First()) : null;

        protected object getValue(Cell cell, Type type)
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
                        return ConvertToTypeWitNullableCheck(cell.CellValue, type);
                    case CellValues.SharedString:
                        var RefId = int.Parse(cell.CellValue.Text);
                        return TypeDescriptor.GetConverter(type).ConvertFromString(_docProvider.sharedStrings[RefId].HasValue() ? _docProvider.sharedStrings[RefId] : string.Empty);
                    case CellValues.String:
                    case CellValues.InlineString:
                        return ConvertToTypeWitNullableCheck(cell.CellValue?.Text, type);
                    default:
                        return ConvertToTypeWitNullableCheck(ConvertTypelessCell(cell), type);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Ошибка преобразования ячеек, адрес ссылки [{cell.CellReference}], искходное значение [{cell.CellValue?.Text}], исходный тип [{cell.DataType?.Value}], стиль [{cell.StyleIndex?.Value}]", ex);
            }
        }

        protected object ConvertTypelessCell(Cell item)
        {
            if (item.StyleIndex != null && _docProvider.styles[item.StyleIndex.Value].NumberFormatId.Value != 0 && item.CellValue != null)
            {
                switch (_docProvider.styles[item.StyleIndex.Value].NumberFormatId.Value)
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
                        return decimalParse(item.CellValue.Text);
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
                        throw new NotImplementedException($"Не реализован обработчик для формата {_docProvider.styles[item.StyleIndex.Value].NumberFormatId.Value}");
                }
            }
            else
            {
                return !string.IsNullOrWhiteSpace(item.CellValue?.Text) ? (decimal?)Convert.ToDecimal(item.CellValue.Text, new System.Globalization.CultureInfo("En")) : null;
            }
        }

        private static decimal decimalParse(string item)
        {
            decimal resul;
            if (decimal.TryParse(item, out resul)) return resul;
            return Convert.ToDecimal(double.Parse(item, new System.Globalization.CultureInfo("En")));
        }
        //=================================================
        #endregion
    }
}