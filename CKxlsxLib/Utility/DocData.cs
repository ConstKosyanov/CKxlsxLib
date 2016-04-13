using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLOC.Utility
{
    public class DocDictionaries : IDisposable
    {
        private SpreadsheetDocument document;
        public string[] sharedStrings;
        public string[] sheetNames;
        public CellFormat[] styles;

        public DocDictionaries(SpreadsheetDocument document)
        {
            this.document = document;
            sharedStrings = document.WorkbookPart.SharedStringTablePart.SharedStringTable.Select(x => x.InnerText).ToArray();
            sheetNames = document.WorkbookPart.Workbook.Sheets.Cast<Sheet>().Select(x => x.Name.Value).ToArray();
            styles = document.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Cast<CellFormat>().ToArray();
        }

        #region Disposing
        //=================================================
        ~DocDictionaries()
        {
            Dispose(false);
        }

        public void Dispose() => Dispose(true);

        void Dispose(bool Disposing)
        {
            if (Disposing)
            {
                document.Close();
            }
        }
        //=================================================
        #endregion
    }
}
