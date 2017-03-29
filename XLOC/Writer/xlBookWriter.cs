using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;
using XLOC.Book;

namespace XLOC.Writer
{
    class XlBookWriter : XlWriter
    {
        #region Variables
        //=================================================
        XlBook _xlBook;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        internal XlBookWriter(XlBook book) => _xlBook = book;
        //=================================================
        #endregion

        #region Private
        //=================================================
        protected override IEnumerable<Sheet> GetSheets()
        {
            foreach (XlSheet item in _xlBook.Sheets)
                yield return new Sheet() { SheetId = item.Id, Name = item.Name };
        }

        protected override IEnumerable<Row> GetRows(uint SheetId) => _xlBook.Sheets.Single(x => x.Id == SheetId)
            .Cells.Select(x => x.Row).Distinct()
            .Select(x => new Row() { RowIndex = new UInt32Value((uint)x) });

        protected override IEnumerable<Cell> GetCellsInRow(UInt32Value SheetId, UInt32Value RowId)
        {
            foreach (XlCell item in _xlBook.Sheets.Single(x => x.Id == SheetId.Value).Cells.Where(x => x.Row == RowId.Value))
                yield return CovertCell(new Cell() { CellReference = item.Reference }, item.Value, item.Type);
        }
        //=================================================
        #endregion
    }
}