using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using qXlsxLib.Utility;
using qXlsxLib.Utility.Events;
using DocumentFormat.OpenXml.Spreadsheet;
using qXlsxLib.Utility.Extensions;
using System.ComponentModel;
using System.Reflection;

namespace qXlsxLib.Reader
{
    class newxlArrayReader
    {
        #region Var
        //=================================================
        qXlsxConfiguration _config;
        DocDictionaries docProvider;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        public newxlArrayReader(qXlsxConfiguration configuration)
        {
            _config = configuration;
            initEvents(configuration);
        }

        void initEvents(qXlsxConfiguration configuration)
        {
            OnValidationFailure += configuration.ValidationFailureEvent;
            OnCellReadingError += configuration.CellReadingErrorEvent;
        }

        //=================================================
        #endregion

        #region Events
        //=================================================
        public event EventHandler<SheetValidationErrorEventArgs> OnValidationFailure;
        public event EventHandler<CellReadingErrorEventArgs> OnCellReadingError;
        //=================================================
        #endregion

        #region Static
        //=================================================
        static T GetAttribute<T>(PropertyInfo x) where T : Attribute => (T)Attribute.GetCustomAttribute(x, typeof(T));
        static bool isRequired(PropertyInfo x) => GetAttribute<xlFieldAttribute>(x).IsRequired;
        static bool isNullable(PropertyInfo x) => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>);
        static IEnumerable<Cell> GetCaptionCells(WorksheetPart sheet) => sheet.Worksheet.GetFirstChild<SheetData>().Descendants<Row>().First().Descendants<Cell>().Where(x => x.CellValue != null);
        static string[][] GetClassCaptions<T>(bool requiredOnly = false) => typeof(T).GetProperties()
            .Where(x => Attribute.IsDefined(x, typeof(xlFieldAttribute)) && !(requiredOnly && (isNullable(x) || !isRequired(x))))
            .Select(x => GetAttribute<xlFieldAttribute>(x).Captions.Union(new string[] { x.Name }).ToArray()).ToArray();
        static xlContentType ToxlContentType(CellValues? local)
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
        //=================================================
        #endregion

        #region Private
        //=================================================

        #region Sheet
        //=================================================
        CellInfo ReadCell(Cell item)
        {
            string Reference = string.Empty;
            object Value = null;
            int? RefId = null;

            Reference = item.CellReference.Value;
            xlContentType? Type = ToxlContentType(item.DataType?.Value);
            try
            {
                switch (Type)
                {
                    case xlContentType.Boolean:
                        throw new NotImplementedException(string.Format("Преобразование для типа {0} не реализовано", xlContentType.Boolean));
                    case xlContentType.Date:
                        throw new NotImplementedException(string.Format("Преобразование для типа {0} не реализовано", xlContentType.Date));
                    case xlContentType.Double:
                        Value = item.CellValue != null ? (decimal?)Convert.ToDecimal(item.CellValue.Text, new System.Globalization.CultureInfo("En")) : null;
                        break;
                    case xlContentType.SharedString:
                        RefId = int.Parse(item.CellValue.Text);
                        Value = docProvider.sharedStrings[RefId.Value].HasValue() ? docProvider.sharedStrings[RefId.Value] : string.Empty;
                        break;
                    case xlContentType.String:
                        Value = (item.CellValue == null ? null : item.CellValue.Text);
                        break;
                    default:
                        Value = ConvertTypelessCell(item);
                        break;
                }
                return new CellInfo(Reference, Type, Value, RefId);
            }
            catch (Exception ex)
            {
                throw new FormatException(string.Format("Ошибка конвертирования ячейки при чтении\nисходное значение: {0}\nконечный тип: {1}", item.CellValue.Text, Type), ex);
            }
        }
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
                        throw new NotImplementedException(string.Format("Не реализован обработчик для формата {0}", docProvider.styles[item.StyleIndex.Value].NumberFormatId.Value));
                }
            }
            else
            {
                return (item.CellValue != null && !string.IsNullOrWhiteSpace(item.CellValue.Text)) ? (decimal?)Convert.ToDecimal(item.CellValue.Text, new System.Globalization.CultureInfo("En")) : null;
            }
        }
        bool isValid<T>(WorksheetPart sheet)
        {
            try
            {
                var captions = GetClassCaptions<T>(true);
                var cells = GetCaptionCells(sheet).Select(x => ReadCell(x).Value.ToString()).ToArray();
                var missingFields = captions.Where(x => x.All(y => !cells.Any(z => string.Equals(y, z, StringComparison.CurrentCultureIgnoreCase)))).Select(x => x.First());
                if (missingFields.Any()) OnValidationFailure(this, new SheetValidationErrorEventArgs(missingFields));
                return !missingFields.Any();
            }
            catch { return false; }
        }
        Map<T> GetMap<T>(WorksheetPart sheet)
        {
            var result = new Map<T>(GetCaptionCells(sheet).ToDictionary(x => x.CellReference.Value, x => ReadCell(x).Value.ToString()));
            foreach (var p in typeof(T).GetProperties().Where(x => Attribute.IsDefined(x, typeof(xlFieldAttribute))).Select(x => new { Propery = x, Attribute = GetAttribute<xlFieldAttribute>(x) }))
            {
                try
                {
                    var cell = cells.SingleOrDefault(x => p.Attribute.Captions.Any(y => string.Equals(x.Value, y, StringComparison.CurrentCultureIgnoreCase)) || string.Equals(x.Value, p.Propery.Name, StringComparison.CurrentCultureIgnoreCase));
                    if (cell.Key.HasValue())
                        result[cell.Key.rMatch(@"^[A-Z]+")] = new MapItem { Property = p.Propery, Attribute = p.Attribute };
                }
                catch (InvalidOperationException ex)
                {
                    throw new InvalidOperationException("Дублирование заголовков", ex);
                }
            }
            return result;
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

                }
                catch (Exception ex)
                {
                    OnCellReadingError(this, new CellReadingErrorEventArgs
                    {
                        Reference = cell.CellReference,
                        SourceType = cell.DataType?.Value ?? null,
                        OutputType = map[cell.CellReference.Value.rMatch("^[A-Z]+")].Attribute.ContentType,
                        Value = cell.CellValue.Text,
                        Exception = ex
                    });
                    throw;
                }
            }
            return result;
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
                    if (!isValid<T>(sheet)) continue;
                    var map = GetMap<T>(sheet);

                    foreach (var row in sheet.Worksheet.Descendants<Row>().Skip(1))
                    {
                        T tmp;
                        try { tmp = RowToObject<T>(row, map); }
                        catch { if (_config.ContinueOnRowReadingError) continue; else throw; }
                        yield return tmp;
                    }
                }
            }
        }
        //=================================================
        #endregion
    }
}