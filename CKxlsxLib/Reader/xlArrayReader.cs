//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;
//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.ComponentModel;
//using qXlsxLib.Utility;
//using qXlsxLib.Utility.Events;
//using qXlsxLib.Utility.Extensions;

//namespace qXlsxLib.Reader
//{
//    internal class xlArrayReader : xlReader
//    {
//        qXlsxConfiguration _config;
//        #region Constructor
//        //=================================================
//        internal xlArrayReader() : base(null) { }
//        public xlArrayReader(SpreadsheetDocument document) : base(document) { }
//        //=================================================
//        #endregion

//        #region private
//        //=================================================
//        static IEnumerable<Cell> GetCaptionCells(WorksheetPart sheet) => sheet.Worksheet.GetFirstChild<SheetData>().Descendants<Row>().First().Descendants<Cell>().Where(x => x.CellValue != null);

//        static bool isRequired(System.Reflection.PropertyInfo x) => GetAttribute<xlFieldAttribute>(x).IsRequired;

//        private static T GetAttribute<T>(System.Reflection.PropertyInfo x) where T : Attribute => (T)Attribute.GetCustomAttribute(x, typeof(T));

//        static bool isNullable(System.Reflection.PropertyInfo x) => (x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>));
//        protected string[][] GetClassCaptions<T>(bool requiredOnly = false) =>
//            typeof(T).GetProperties()
//            .Where(x => Attribute.IsDefined(x, typeof(xlFieldAttribute)) && !(requiredOnly && (isNullable(x) || !isRequired(x))))
//            .Select(x => ((xlFieldAttribute)Attribute.GetCustomAttribute(x, typeof(xlFieldAttribute))).Captions.Union(new string[] { x.Name }).ToArray())
//            .ToArray();

//        bool isValid<T>(WorksheetPart sheet)
//        {
//            try
//            {
//                var captions = GetClassCaptions<T>(true);
//                var cells = GetCaptionCells(sheet).Select(x => ReadCell(x).Value.ToString()).ToArray();
//                var missingFields = captions.Where(x => x.All(y => !cells.Any(z => string.Equals(y, z, StringComparison.CurrentCultureIgnoreCase)))).Select(x => x.First());
//                if (missingFields.Any()) OnValidationFailure(this, new SheetValidationErrorEventArgs(missingFields));
//                return !missingFields.Any();
//            }
//            catch
//            {
//                return false;
//            }
//        }

//        Dictionary<string, MapItem> GetMap<T>(WorksheetPart sheet)
//        {
//            var result = new Dictionary<string, MapItem>();
//            var cells = GetCaptionCells(sheet).ToDictionary(x => x.CellReference.Value, x => ReadCell(x).Value.ToString()); ;
//            foreach (var p in typeof(T).GetProperties().Where(x => Attribute.IsDefined(x, typeof(xlFieldAttribute))).Select(x => new { Propery = x, Attribute = (xlFieldAttribute)Attribute.GetCustomAttribute(x, typeof(xlFieldAttribute)) }))
//            {
//                try
//                {
//                    var cell = cells.SingleOrDefault(x => p.Attribute.Captions.Any(y => string.Equals(x.Value, y, StringComparison.CurrentCultureIgnoreCase)) || string.Equals(x.Value, p.Propery.Name, StringComparison.CurrentCultureIgnoreCase));
//                    if (cell.Key.HasValue())
//                        result[cell.Key.rReplace(@"^([A-Z]+)\d+", @"$1.+")] = new MapItem { Property = p.Propery, Attribute = ((xlFieldAttribute)Attribute.GetCustomAttribute(p.Propery, typeof(xlFieldAttribute))) };
//                }
//                catch (InvalidOperationException ex)
//                {
//                    throw new InvalidOperationException("Дублирование заголовков", ex);
//                }
//            }
//            return result;
//        }

//        private object ConvertValue(CellInfo local, xlContentType destinationType, bool Nullable)
//        {
//            try
//            {
//                if ((local.Value == null) && Nullable)
//                    return null;

//                switch (destinationType)
//                {
//                    case xlContentType.Boolean:
//                        return local.ContentType == destinationType ? local.Value : Convert.ToBoolean(local.Value, new System.Globalization.CultureInfo("En"));
//                    case xlContentType.Integer:
//                        return local.ContentType == destinationType ? local.Value : Convert.ToInt32(local.Value, new System.Globalization.CultureInfo("En"));
//                    case xlContentType.Double:
//                        return local.ContentType == destinationType ? local.Value : Convert.ToDecimal(local.Value.ToString().rReplace("(?<=\\d)[\\.,](?=\\d)", ","));
//                    case xlContentType.SharedString:
//                        return local.ContentType == destinationType ? local.Value : Convert.ToString(local.Value, new System.Globalization.CultureInfo("En"));
//                    case xlContentType.String:
//                        return local.ContentType == destinationType ? local.Value : Convert.ToString(local.Value, new System.Globalization.CultureInfo("En"));
//                    case xlContentType.Date:
//                        return local.ContentType == destinationType ? local.Value : Convert.ToDateTime(local.Value);
//                    default:
//                        throw new ArgumentException("Произошла ошибка при конвертировании ячеек, тип ячейки не распознан");
//                }
//            }
//            catch (Exception ex)
//            {
//                if (Nullable) return null;
//                throw new InvalidCastException("Ошбика преобразования ячеек", ex);
//            }
//        }
//        //=================================================
//        #endregion

//        #region Events
//        //=================================================
//        public event EventHandler<SheetValidationErrorEventArgs> OnValidationFailure;
//        public event EventHandler<CellReadingErrorEventArgs> OnCellReadingError;
//        //=================================================
//        #endregion

//        #region Methods
//        //=================================================
//        public IEnumerable<T> ReadToEnumerable<T>() where T : IxlCompatible, new()
//        {
//            var sheets = _config.Sheets == null ? doc.WorkbookPart.Workbook.Sheets.Cast<Sheet>() : doc.WorkbookPart.Workbook.Sheets.Cast<Sheet>().Where(x => _config.Sheets.Contains(x.SheetId.Value)).ToArray();
//            foreach (var sheet in doc.WorkbookPart.WorksheetParts.Where(x => sheets.Select(y => y.Id.Value).Contains(doc.WorkbookPart.GetIdOfPart(x))))
//            {
//                if (!isValid<T>(sheet)) continue;
//                var map = GetMap<T>(sheet);

//                foreach (var row in sheet.Worksheet.Descendants<Row>().Skip(1))
//                {
//                    try
//                    {
//                        foreach (var cell in row.Descendants<Cell>().Where(x => map.Keys.Any(y => x.CellReference.Value.rIsMatch(y))))
//                        {

//                        }
//                    }
//                    catch (Exception ex)
//                    {
//                        throw ex;
//                    }
//                }

//                foreach (var row in sheet.Worksheet.Descendants<Row>().Skip(1))
//                {
//                    var tmpRes = new T();
//                    try
//                    {
//                        foreach (var cell in row.Descendants<Cell>().Where(x => map.Keys.Any(y => x.CellReference.Value.rIsMatch(y))).Select(x => new { Reference = x.CellReference.Value.rReplace(@"^([A-Z]+)\d+", @"$1.+"), Value = ReadCell(x) }))
//                        {
//                            try
//                            {
//                                map[cell.Reference].Property.SetValue(tmpRes, ConvertValue(cell.Value, map[cell.Reference].Attribute.ContentType, (map[cell.Reference].Property.PropertyType.IsGenericType && map[cell.Reference].Property.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>))));
//                            }
//                            catch (Exception ex)
//                            {
//                                OnCellReadingError(this, new CellReadingErrorEventArgs { Reference = cell.Reference, SourceType = cell.Value.ContentType, OutputType = map[cell.Reference].Attribute.ContentType, Value = cell.Value.Value, Exception = ex });
//                                throw;
//                            }
//                        }
//                    }
//                    catch
//                    {
//                        continue;
//                    }
//                    yield return tmpRes;
//                }
//            }
//        }
//        //=================================================
//        #endregion
//    }
//}