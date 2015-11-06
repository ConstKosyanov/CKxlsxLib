using CKxlsxLib.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CKxlsxLib.Reader
{
    public class xlReader : IDisposable
    {
        #region Variables
        //=================================================
        protected SpreadsheetDocument doc;
        protected string[] sharedStrings;
        protected string[] sheetNames;
        CellFormat[] styles;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        protected xlReader(SpreadsheetDocument document)
        {
            doc = document;
            sharedStrings = doc.WorkbookPart.SharedStringTablePart.SharedStringTable.Select(x => x.InnerText).ToArray();
            sheetNames = doc.WorkbookPart.Workbook.Sheets.Cast<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Select(x => x.Name.Value).ToArray();
            styles = doc.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Cast<CellFormat>().ToArray();
        }

        #region Destructor
        //=================================================
        ~xlReader()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
        }

        void Dispose(bool Disposing)
        {
            if (Disposing)
            {
                doc.Close();
            }
        }
        //=================================================
        #endregion
        //=================================================
        #endregion

        #region Factory
        //=================================================
        public static xlReader FromFile(string path)
        {
            try
            {
                return xlReader.FromStream(new MemoryStream(File.ReadAllBytes(path)));
            }
            catch (Exception ex)
            {
                throw new IOException(string.Format("Не удалось открыть файл {0}", path), ex);
            }
        }

        public static xlReader FromStream(Stream stream)
        {
            return new xlReader(SpreadsheetDocument.Open(stream, false));
        }

        public static xlReader FromBuffer(byte[] buffer)
        {
            return new xlReader(SpreadsheetDocument.Open(new MemoryStream(buffer), false));
        }
        //=================================================
        #endregion

        #region private
        //=================================================
        protected object ConvertTypelessCell(Cell item)
        {
            if (item.StyleIndex != null && styles[item.StyleIndex.Value].NumberFormatId.Value != 0 && item.CellValue != null)
            {
                switch (styles[item.StyleIndex.Value].NumberFormatId.Value)
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
                        throw new NotImplementedException(string.Format("Не реализован обработчик для формата {0}", styles[item.StyleIndex.Value].NumberFormatId.Value));
                }
            }
            else
            {
                return item.CellValue != null ? (decimal?)Convert.ToDecimal(item.CellValue.Text, new System.Globalization.CultureInfo("En")) : null;
            }
        }

        protected Tuple<string, xlContentType?, object, int?> ReadCell(Cell item)
        {
            string Reference = string.Empty;
            object Value = null;
            int? RefId = null;

            Reference = item.CellReference.Value;
            xlContentType? Type = (item.DataType != null ? (DocumentFormat.OpenXml.Spreadsheet.CellValues?)item.DataType.Value : null).ToxlContentType();
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
                        Value = sharedStrings[RefId.Value].HasValue() ? sharedStrings[RefId.Value] : string.Empty;
                        break;
                    case xlContentType.String:
                        Value = (item.CellValue == null ? null : item.CellValue.Text);
                        break;
                    //throw new NotImplementedException(string.Format("Преобразование для типа {0} не реализовано", xlContentType.String));
                    default:
                        Value = ConvertTypelessCell(item);
                        break;
                }
                return new Tuple<string, xlContentType?, object, int?>(Reference, Type, Value, RefId);
            }
            catch (Exception ex)
            {
                throw new FormatException(string.Format("Ошибка конвертирования ячейки при чтении\nисходное значение: {0}\nконечный тип: {1}", item.CellValue.Text, Type), ex);
            }
        }
        //=================================================
        #endregion

        #region Methods
        //=================================================
        public virtual xlBook ReadToBook()
        {
            return new xlBookReader(doc).ReadToBook();
        }

        public virtual IEnumerable<T> ReadToEnumerable<T>(uint[] SheetIDs = null, EventHandler<CKxlsxLibEventArgs> OnValidationFailure = null) where T : IxlCompatible, new()
        {
            var xlArrayReader = new xlArrayReader(doc);
            if (OnValidationFailure != null)
                xlArrayReader.OnValidationFailure += OnValidationFailure;
            foreach (var item in xlArrayReader.ReadToEnumerable<T>(SheetIDs, OnValidationFailure))
            {
                yield return item;
            }
        }

        public virtual T[] ReadToArray<T>(uint[] SheetIDs = null, EventHandler<CKxlsxLibEventArgs> OnValidationFailure = null) where T : IxlCompatible, new()
        {
            return this.ReadToEnumerable<T>(SheetIDs, OnValidationFailure).ToArray();
        }
        //=================================================
        #endregion
    }
}