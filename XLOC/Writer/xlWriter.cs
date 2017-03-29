using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XLOC.Book;
using XLOC.Utility;
using XLOC.Utility.Extensions;

namespace XLOC.Writer
{
    public abstract class XlWriter
    {
        #region Variables
        //=================================================
        protected Dictionary<string, int> sst;
        protected int sstNext = 0;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        public XlWriter() => sst = new Dictionary<string, int>();
        //=================================================
        #endregion

        #region private
        //=================================================
        void intiWorkbook(SpreadsheetDocument document)
        {
            document.AddWorkbookPart();
            document.WorkbookPart.Workbook = new Workbook();
            document.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            initStylesPart(document.WorkbookPart.AddNewPart<WorkbookStylesPart>());
        }

        WorksheetPart initWorksheetPart(SpreadsheetDocument document, Sheet sheet)
        {
            WorksheetPart wsp = document.WorkbookPart.AddNewPart<WorksheetPart>();
            wsp.Worksheet = new Worksheet(new SheetData());
            sheet.Id = document.WorkbookPart.GetIdOfPart(wsp);
            document.WorkbookPart.Workbook.Sheets.Append(sheet);
            return wsp;
        }

        void initStylesPart(WorkbookStylesPart stylesPart)
        {
            #region Old
            //=================================================
            stylesPart.Stylesheet = new Stylesheet() { Fonts = new Fonts() };

            #region Fonts
            //=================================================
            stylesPart.Stylesheet.Fonts.AppendChild(new Font());
            //=================================================
            #endregion

            #region Fills
            //=================================================
            stylesPart.Stylesheet.Fills = new Fills();
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } });
            //=================================================
            #endregion

            #region Borders
            //=================================================
            stylesPart.Stylesheet.Borders = new Borders();
            stylesPart.Stylesheet.Borders.AppendChild(new Border());
            //=================================================
            #endregion

            #region CellStyles
            //=================================================         
            stylesPart.Stylesheet.CellStyles = new CellStyles();
            stylesPart.Stylesheet.CellStyles.Append(new CellStyle() { FormatId = 0 });
            stylesPart.Stylesheet.CellStyles.Append(new CellStyle() { FormatId = 1 });
            stylesPart.Stylesheet.CellStyles.Append(new CellStyle() { FormatId = 2 });
            //=================================================
            #endregion

            #region CellStyleFormats
            //=================================================
            stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
            stylesPart.Stylesheet.CellStyleFormats.Append(new CellFormat());
            stylesPart.Stylesheet.CellStyleFormats.Append(new CellFormat());
            stylesPart.Stylesheet.CellStyleFormats.Append(new CellFormat());
            //=================================================
            #endregion

            #region CellFormats
            //=================================================
            stylesPart.Stylesheet.CellFormats = new CellFormats();
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat()
            {
                FormatId = 1,
                FontId = 0,
                BorderId = 0,
                FillId = 0,
                NumberFormatId = 14,
            });
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat()
            {
                FormatId = 2,
                FontId = 0,
                BorderId = 0,
                FillId = 0,
                NumberFormatId = 2,
            });
            //=================================================
            #endregion

            stylesPart.Stylesheet.Save();
            //=================================================
            #endregion
        }

        void intiSharedStringTablePart(SharedStringTablePart sharedStringTablePart)
        {
            sharedStringTablePart.SharedStringTable = new SharedStringTable();
            sharedStringTablePart.SharedStringTable.Append(sst.Keys.Select(x => new SharedStringItem(new Text(x))));
            sharedStringTablePart.SharedStringTable.Save();
        }

        void fillSheet(WorksheetPart p, Sheet x)
        {
            foreach (Row item in GetRows(x.SheetId))
            {
                item.Append(GetCellsInRow(x.SheetId, item.RowIndex));
                p.Worksheet.GetFirstChild<SheetData>().Append(item);
            }
        }

        protected int getSharedStringId(string Key) => sst.ContainsKey(Key) ? sst[Key] : sst[Key] = sstNext++;

        protected Cell CovertCell(Cell cell, object Value, XlContentType? Type)
        {
            if (Value == null)
                return cell;

            try
            {
                switch (Type)
                {
                    case XlContentType.Boolean:
                        throw new NotImplementedException(string.Format("Не реализован обработчик записи ячеек с типом {0}", XlContentType.Boolean));
                    case XlContentType.Integer:
                        cell.CellValue = new CellValue(Value.ToString());
                        break;
                    case XlContentType.Double:
                        cell.CellValue = new CellValue(Convert.ToString(Value, new System.Globalization.CultureInfo("En")));
                        cell.DataType = null;
                        cell.StyleIndex = 2;
                        break;
                    case XlContentType.SharedString:
                        cell.CellValue = new CellValue(getSharedStringId(Value.ToString()).ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        break;
                    case XlContentType.String:
                        throw new NotImplementedException(string.Format("Не реализован обработчик записи ячеек с типом {0}", XlContentType.String));
                    case XlContentType.Date:
                        cell.CellValue = new CellValue(((DateTime)Value).ToOADate().ToString());
                        cell.DataType = null;
                        cell.StyleIndex = 1;
                        break;
                    default:
                        cell.CellValue = new CellValue(Value.ToString());
                        break;
                }
                return cell;
            }
            catch (Exception ex)
            {
                throw new Exception("Ошибка преобразования ячейки", ex);
            }
        }

        protected abstract IEnumerable<Sheet> GetSheets();
        protected abstract IEnumerable<Row> GetRows(uint SheetId);
        protected abstract IEnumerable<Cell> GetCellsInRow(UInt32Value SheetId, UInt32Value RowId);
        //=================================================
        #endregion

        #region Methods
        //=================================================
        public ValidationErrorInfo[] SaveToFile(string path)
        {
            try
            {
                using (FileStream file = File.Open(path, FileMode.OpenOrCreate))
                {
                    return SaveToStream(file);
                }
            }
            catch (Exception ex)
            {
                throw new IOException(string.Format("Не удалось сохранить файл {0}", path), ex);
            }
        }

        public ValidationErrorInfo[] SaveToStream(Stream stream)
        {
            sst = new Dictionary<string, int>();
            using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                intiWorkbook(doc);
                GetSheets().ForEach((x) => fillSheet(initWorksheetPart(doc, x), x));
                intiSharedStringTablePart(doc.WorkbookPart.AddNewPart<SharedStringTablePart>());
                doc.WorkbookPart.Workbook.Save();

                OpenXmlValidator validator = new OpenXmlValidator();
                return validator.Validate(doc).ToArray();
            }
        }

        public ValidationErrorInfo[] SaveToBuffer(out byte[] result)
        {
            MemoryStream stream = new MemoryStream();
            ValidationErrorInfo[] res = SaveToStream(stream);
            result = stream.ToArray();
            return res;
        }
        //=================================================
        #endregion
    }
}