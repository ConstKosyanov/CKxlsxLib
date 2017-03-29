using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLOC.Utility;
using XLOC.Utility.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

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

        #region private
        //=================================================

        #region Factory
        //=================================================
        static Sheet createSheet() => new Sheet()
        {
            SheetId = new UInt32Value(_defId),
            Name = new StringValue(new string(typeof(T).Name.Take(_SheetNameMaxLength).ToArray()))
        };

        static Row createRow(uint i) => new Row()
        {
            RowIndex = new UInt32Value(i),
            Spans = new ListValue<StringValue>(new StringValue[] { new StringValue($"1:{typeof(T).GetProperties().Count()}") })
        };

        static Cell createCell(int index, int ColCounter) => new Cell()
        {
            CellReference = new StringValue(Book.XlCell.GetReference(ColCounter, index))
        };
        //=================================================
        #endregion

        static XlContentType getContentType(System.Reflection.PropertyInfo property) => ((XlFieldAttribute)Attribute.GetCustomAttribute(property, typeof(XlFieldAttribute))).ContentType;

        protected override IEnumerable<Sheet> GetSheets() => new[] { createSheet() };

        protected override IEnumerable<Row> GetRows(uint SheetId)
        {
            for (uint i = 1; i <= items.Count() + 1; i++)
                yield return createRow(i);
        }

        protected override IEnumerable<Cell> GetCellsInRow(UInt32Value SheetId, UInt32Value RowId) => RowId.Value == 1
            ? getCaptions()
            : getValuesFrom((int)RowId.Value);

        IEnumerable<Cell> getCaptions()
        {
            int ColCounter = 1;
            foreach (PropertyInfo property in typeof(T).GetProperties().Where(x => Attribute.IsDefined(x, typeof(XlFieldAttribute))))
            {
                Cell cell = createCell(1, ColCounter++);
                cell.CellValue = new CellValue(getSharedStringId(((XlFieldAttribute)Attribute.GetCustomAttribute(property, typeof(XlFieldAttribute))).Captions.First()).ToString());
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                yield return cell;
            }
        }

        static IEnumerable<PropertyInfo> getProperties() => typeof(T).GetProperties().Where(x => Attribute.IsDefined(x, typeof(XlFieldAttribute)));

        IEnumerable<Cell> getValuesFrom(int index)
        {
            T item = items[index - 2];
            int ColCounter = 0;
            foreach (PropertyInfo property in getProperties())
            {
                ColCounter++;
                yield return property.GetValue(item) != null
                    ? CovertCell(createCell(index, ColCounter), property.GetValue(item), getContentType(property))
                    : null;
            }
        }
        //=================================================
        #endregion
    }
}