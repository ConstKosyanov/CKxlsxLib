using System;
using System.Collections.Generic;
using System.Linq;
using XLOC.Utility;
using XLOC.Utility.Events;
using XLOC.Utility.Extensions;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using XLOC.Utility.Exceptions;

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

        #region Map
        //=================================================
        int getSkip() => _config.SkipMode == SkipModeEnum.None ? 1 : _config.SkipCount ?? 0;

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

        Dictionary<string, string> ToDictionary(Row row) => row.Descendants<Cell>().Where(x => x.CellValue != null).ToDictionary(x => x.CellReference.Value, x => getValue(x, typeof(string)).ToString());

        Map<T> AutoMap<T>(WorksheetPart sheet)
        {
            Map<T> result = null;
            _config.SkipCount = 0;
            foreach (var item in getRows(sheet.Worksheet.GetFirstChild<SheetData>()))
            {
                result = new Map<T>(ToDictionary(item));
                _config.SkipCount++;
                if (result.IsValid) return result;
            }
            return result;
        }

        Map<T> GetMap<T>(WorksheetPart sheet) where T : new()
        {
            switch (_config.SkipMode)
            {
                case SkipModeEnum.None: return new Map<T>(sheet.GetCaptionCells().ToDictionary(x => x.CellReference.Value, x => getValue(x, typeof(string)).ToString()));
                case SkipModeEnum.Manual: return new Map<T>(sheet.GetCaptionCells(_config.SkipCount.Value).ToDictionary(x => x.CellReference.Value, x => getValue(x, typeof(string)).ToString()));
                case SkipModeEnum.Auto: return AutoMap<T>(sheet);
                default: throw new NotImplementedException("Default switch case not implemented");
            }
        }
        //=================================================
        #endregion

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
                    throw new CellReadingException(ex)
                    {
                        Reference = cell.CellReference,
                        SourceType = cell.DataType?.Value ?? null,
                        OutputType = map[cell.CellReference.Value.rMatch("^[A-Z]+")].Attribute.ContentType,
                        Value = cell.CellValue?.Text,
                    };
                }
            }
            return result;
        }

        string getWorksheetPartId(WorksheetPart wSheet) => _config.Document.WorkbookPart.GetIdOfPart(wSheet);

        SheetIdentifier createSheetIdentifier(WorksheetPart wSheet) => new SheetIdentifier(_config.Document.WorkbookPart.Workbook.Sheets.Cast<Sheet>().Single(x => x.Id == getWorksheetPartId(wSheet)));

        IEnumerable<Sheet> getFilteredSheets() => _config.Sheets == null
            ? _config.Document.WorkbookPart.Workbook.Sheets.Cast<Sheet>()
            : _config.Document.WorkbookPart.Workbook.Sheets.Cast<Sheet>().Where(x => _config.Sheets.Contains(x.SheetId.Value)).ToArray();

        IEnumerable<WorksheetPart> getFilteredWorksheetParts() => _config.Document.WorkbookPart.WorksheetParts.Where(x => getFilteredSheets().Select(y => y.Id.Value).Contains(getWorksheetPartId(x)));

        IEnumerable<T> readSheet<T>(WorksheetPart wSheet) where T : new()
        {
            Map<T> map = GetMap<T>(wSheet);
            var Sheet = createSheetIdentifier(wSheet);

            if (!map.IsValid)
            {
                OnValidationFailure(this, new SheetValidationErrorEventArgs(map.MissingFields, map.Exceptioins, Sheet));
                yield break;
            }

            foreach (var row in wSheet.Worksheet.Descendants<Row>().Skip(getSkip()))
            {
                T tmp;
                try { tmp = RowToObject<T>(row, map); }
                catch (CellReadingException ex)
                {
                    cellErrorEventCaller(new CellReadingErrorEventArgs
                    {
                        Reference = ex.Reference,
                        SourceType = ex.SourceType,
                        OutputType = ex.OutputType,
                        Value = ex.Value,
                        Row = row.RowIndex.Value,
                        Sheet = Sheet,
                    });

                    if (_config.ContinueOnRowReadingError)
                        continue;
                    else
                        throw;
                }
                yield return tmp;
            }
        }
        //=================================================
        #endregion

        #region Methods
        //=================================================
        internal IEnumerable<T> ReadToEnumerable<T>() where T : new() => getFilteredWorksheetParts().SelectMany(readSheet<T>);

        internal IEnumerable<IGrouping<SheetIdentifier, T>> ReadToGroup<T>() where T : new() => getFilteredWorksheetParts().Select(x => XLOCSheet<T>.Create(createSheetIdentifier(x), readSheet<T>(x)));
        //=================================================
        #endregion
    }
}