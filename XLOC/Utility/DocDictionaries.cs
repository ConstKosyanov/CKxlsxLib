using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLOC.Utility
{
    public sealed class DocDictionaries : IDisposable
    {
        SpreadsheetDocument document;
        bool _autoDispose;
        public string[] sharedStrings;
        public string[] sheetNames;
        public CellFormat[] styles;

        public DocDictionaries(SpreadsheetDocument document, bool AutoDispose)
        {
            _autoDispose = AutoDispose;
            this.document = document;
            sharedStrings = document.WorkbookPart.SharedStringTablePart.SharedStringTable.Select(x => x.InnerText).ToArray();
            sheetNames = document.WorkbookPart.Workbook.Sheets.Cast<Sheet>().Select(x => x.Name.Value).ToArray();
            styles = document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Cast<CellFormat>().ToArray();
        }

        #region Disposing
        //=================================================
        ~DocDictionaries()
        {
            dispose(false);
        }

        public void Dispose() => dispose(true);

        void dispose(bool Disposing)
        {
            if (Disposing)
            {
                if (_autoDispose)
                {
                    document.Close();
                }
            }
        }
        //=================================================
        #endregion
    }
}