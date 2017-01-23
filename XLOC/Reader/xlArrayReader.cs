using System;
using System.Collections.Generic;
using System.Linq;
using XLOC.Utility;
using XLOC.Utility.Events;
using XLOC.Utility.Extensions;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace XLOC.Reader
{
    class xlArrayReader : xlReader
    {
        #region Events
        //=================================================
        public event EventHandler<SheetValidationErrorEventArgs> OnValidationFailure;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        public xlArrayReader(XLOCConfiguration configuration) : base(configuration)
        {
            OnValidationFailure += (s, e) => { };
            OnValidationFailure += _config.ValidationFailureEvent;
        }
        //=================================================
        #endregion

        #region Private
        //=================================================
        T RowToObject<T>(Row row, Map<T> map) where T : new()
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
                    cellErrorEventCaller(new CellReadingErrorEventArgs
                    {
                        Reference = cell.CellReference,
                        SourceType = cell.DataType?.Value ?? null,
                        OutputType = map[cell.CellReference.Value.rMatch("^[A-Z]+")].Attribute.ContentType,
                        Value = cell.CellValue?.Text,
                        Sheet = map.Sheet,
                        Exception = ex
                    });
                    throw;
                }
            }
            return result;
        }

        int getSkip() => _config.SkipMode == SkipModeEnum.None ? 1 : _config.SkipCount ?? 0;

        Map<T> GetMap<T>(WorksheetPart sheet) where T : new()
        {
            switch (_config.SkipMode)
            {
                case SkipModeEnum.None: return new Map<T>(sheet.GetCaptionCells().ToDictionary(x => x.CellReference.Value, x => getValue(x, typeof(string)).ToString()), getSheetIdentifier(sheet));
                case SkipModeEnum.Manual: return new Map<T>(sheet.GetCaptionCells(_config.SkipCount.Value).ToDictionary(x => x.CellReference.Value, x => getValue(x, typeof(string)).ToString()), getSheetIdentifier(sheet));
                case SkipModeEnum.Auto: return AutoMap<T>(sheet);
                default: throw new NotImplementedException("Default switch case not implemented");
            }
        }

        SheetIdentifier getSheetIdentifier(WorksheetPart wSheet) => new SheetIdentifier(_config.Document.WorkbookPart.Workbook.Sheets.Cast<Sheet>().Single(x => x.Id == getWorksheetPartId(wSheet)));

        string getWorksheetPartId(WorksheetPart wSheet) => _config.Document.WorkbookPart.GetIdOfPart(wSheet);

        Map<T> AutoMap<T>(WorksheetPart sheet)
        {
            Map<T> result = null;
            _config.SkipCount = 0;
            //var enumerator = sheet.Worksheet.GetFirstChild<SheetData>().Descendants<Row>().GetEnumerator();
            //var row = sheet.Worksheet.GetFirstChild<SheetData>().GetFirstChild<Row>();
            foreach (var item in getRows(sheet.Worksheet.GetFirstChild<SheetData>()))
            {
                result = new Map<T>(ToDictionary(item), getSheetIdentifier(sheet));
                _config.SkipCount++;
                if (result.IsValid) return result;
            }
            return result;
        }

        IEnumerable<Row> getRows(SheetData sheet)
        {
            Row tmp = sheet.GetFirstChild<Row>();
            yield return tmp;
            while (tmp != null)
            {
                tmp = tmp.NextSibling<Row>();
                yield return tmp;
            }
        }

        Dictionary<string, string> ToDictionary(Row row)
        {
            return row.Descendants<Cell>().Where(x => x.CellValue != null).ToDictionary(x => x.CellReference.Value, x => getValue(x, typeof(string)).ToString());
        }
        //=================================================
        #endregion

        #region Methods
        //=================================================
        public IEnumerable<T> ReadToEnumerable<T>() where T : new()
        {
            using (_docProvider = new DocDictionaries(_config.Document, _config.AutoDispose))
            {
                var sheets = _config.Sheets == null ? _config.Document.WorkbookPart.Workbook.Sheets.Cast<Sheet>() : _config.Document.WorkbookPart.Workbook.Sheets.Cast<Sheet>().Where(x => _config.Sheets.Contains(x.SheetId.Value)).ToArray();
                foreach (var sheet in _config.Document.WorkbookPart.WorksheetParts.Where(x => sheets.Select(y => y.Id.Value).Contains(_config.Document.WorkbookPart.GetIdOfPart(x))))
                {
                    Map<T> map = GetMap<T>(sheet);

                    if (!map.IsValid)
                    {
                        OnValidationFailure(this, new SheetValidationErrorEventArgs(map.MissingFields, map.Exceptioins, map.Sheet));
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
        //=================================================
        #endregion
    }
}