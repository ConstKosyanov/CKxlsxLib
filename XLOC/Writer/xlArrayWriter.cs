using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLOC.Utility;
using XLOC.Utility.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace XLOC.Writer
{
    internal class XlArrayWriter<T> : XlWriter
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
        public XlArrayWriter(IEnumerable<T> items) => this.items = items.ToArray();
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
                return getCaptions();
            }
            else
            {
                return getValuesFrom((int)RowId.Value);
            }
        }

        IEnumerable<Cell> getValuesFrom(int index)
        {
            T item = items[index - 2];
            int ColCounter = 1;
            foreach (System.Reflection.PropertyInfo property in typeof(T).GetProperties().Where(x => Attribute.IsDefined(x, typeof(XlFieldAttribute))))
            {
                yield return property.GetValue(item) != null ? CovertCell(new Cell() { CellReference = new StringValue(Book.XlCell.GetReference(ColCounter, index)) }, property.GetValue(item), ((XlFieldAttribute)Attribute.GetCustomAttribute(property, typeof(XlFieldAttribute))).ContentType) : null;
                ColCounter++;
            }
            yield break;
        }

        IEnumerable<Cell> getCaptions()
        {
            int ColCounter = 1;
            foreach (System.Reflection.PropertyInfo property in typeof(T).GetProperties().Where(x => Attribute.IsDefined(x, typeof(XlFieldAttribute))))
            {
                var attr = (XlFieldAttribute)Attribute.GetCustomAttribute(property, typeof(XlFieldAttribute));
                var cell = new Cell() { CellReference = new StringValue(Book.XlCell.GetReference(ColCounter++, 1)) };
                cell.CellValue = new CellValue(getSharedStringId(attr.Captions.First()).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                yield return cell;
            }
        }
        //=================================================
        #endregion
    }
}