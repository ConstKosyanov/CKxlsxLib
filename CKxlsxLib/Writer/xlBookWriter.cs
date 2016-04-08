using qXlsxLib.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace qXlsxLib.Writer
{
    class xlBookWriter : xlWriter
    {
        #region Variables
        //=================================================
        xlBook _xlBook;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        internal xlBookWriter(xlBook book)
        {
            _xlBook = book;
        }
        //=================================================
        #endregion

        #region Private
        //=================================================
        protected override IEnumerable<Sheet> GetSheets()
        {
            foreach (var item in _xlBook.Sheets)
            {
                yield return new Sheet() { SheetId = item.Id, Name = item.Name };
            }
        }
        
        protected override IEnumerable<Row> GetRows(uint SheetId)
        {
            return _xlBook.Sheets.Single(x=>x.Id == SheetId).Cells.Select(x => x.Row).Distinct().Select(x => new Row() { RowIndex = new UInt32Value((uint)x) });
        }

        protected override IEnumerable<Cell> GetCellsInRow(UInt32Value SheetId, UInt32Value RowId)
        {
            foreach (var item in _xlBook.Sheets.Single(x => x.Id == SheetId.Value).Cells.Where(x => x.Row == RowId.Value))
            {
                yield return CovertCell(new Cell() { CellReference = item.Reference }, item.Value, item.Type);
            }
        }
        //=================================================
        #endregion
    }
}