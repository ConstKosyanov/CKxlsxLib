using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using qXlsxLib.Utility;
using qXlsxLib.Utility.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace qXlsxLib.Writer
{
    internal class xlArrayWriter<T> : xlWriter
    {
        #region Variables
        //=================================================
        const uint _defId = 1;
        const int _SheetNameMaxLength = 31;
        private T[] items;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        public xlArrayWriter(IEnumerable<T> items)
        {
            this.items = items.ToArray();
        }
        //=================================================
        #endregion

        #region Private
        //=================================================
        protected override IEnumerable<Sheet> GetSheets()
        {
            yield return new Sheet() { SheetId = new UInt32Value(_defId), Name = new StringValue(new string(typeof(T).ToString().rMatch(@"(?<=\.).+?$").Take(_SheetNameMaxLength).ToArray())) };
        }

        protected override IEnumerable<Row> GetRows(uint SheetId)
        {
            Row[] result = new Row[items.Count() + 1];
            for (uint i = 0; i < result.Length; i++)
            {
                result[i] = new Row() { RowIndex = new UInt32Value(i+1), Spans = new ListValue<StringValue>(new StringValue[] { new StringValue(string.Format("1:{0}", typeof(T).GetProperties().Count())) }) };
            }
            return result;
        }

        protected override IEnumerable<Cell> GetCellsInRow(UInt32Value SheetId, UInt32Value RowId)
        {
            if (RowId.Value == 1)
            {
                return GetCaptions();
            }
            else
            {
                return GetValuesFrom((int)RowId.Value);
            }
        }

        private IEnumerable<Cell> GetValuesFrom(int index)
        {
            var item = items[index - 2];
            int ColCounter = 1;
            foreach (var property in typeof(T).GetProperties().Where(x => xlFieldAttribute.IsDefined(x, typeof(xlFieldAttribute))))
            {
                yield return property.GetValue(item) != null ? CovertCell(new Cell() { CellReference = new StringValue(Book.xlCell.GetReference(ColCounter, index)) }, property.GetValue(item), ((xlFieldAttribute)Attribute.GetCustomAttribute(property, typeof(xlFieldAttribute))).ContentType) : null;
                ColCounter++;
            }
            yield break;
        }

        private IEnumerable<Cell> GetCaptions()
        {
            int ColCounter = 1;
            foreach (var property in typeof(T).GetProperties().Where(x => xlFieldAttribute.IsDefined(x, typeof(xlFieldAttribute))))
            {
                var attr = (xlFieldAttribute)Attribute.GetCustomAttribute(property, typeof(xlFieldAttribute));
                var cell = new Cell() { CellReference = new StringValue(Book.xlCell.GetReference(ColCounter++, 1)) };
                cell.CellValue = new CellValue(getSharedStringId(attr.Captions.First()).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                yield return cell;
            }
        }
        //=================================================
        #endregion
    }
}