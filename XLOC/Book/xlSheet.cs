using XLOC.Utility;
using XLOC.Utility.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace XLOC.Book
{
    public class XlSheet
    {
        #region Variables
        //=================================================
        string _name = string.Empty;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        XlSheet(uint Id)
        {
            Cells = new List<XlCell>();
            this.Id = Id;
        }

        internal XlSheet(uint Id, string ListName) : this(Id) => Name = ListName;
        //=================================================
        #endregion

        #region Methods
        //=================================================
        public XlCell AddCell()
        {
            var result = new XlCell();
            Cells.Add(result);
            return result;
        }

        public XlCell AddCell(object cellValue, int aRow, int aCol, XlContentType aType)
        {
            XlCell result = AddCell();
            result.Value = cellValue;
            result.Row = aRow;
            result.Col = aCol;
            result.Type = aType;
            return result;
        }

        public XlCell AddCell(object cellValue, string aReference, XlContentType aType)
        {
            XlCell result = AddCell();
            result.Value = cellValue;
            result.Reference = aReference;
            result.Type = aType;
            return result;
        }
        //=================================================
        #endregion

        #region Properties
        //=================================================
        public string Name
        {
            get => _name;
            set
            {
                if (!value.rIsMatch(@"^[^\:\\\/\?\*\[\]]+$"))
                    throw new ArgumentException("Имя книги содержит недопустимые символы");
                _name = value;
            }
        }

        public IEnumerable<XlCell> this[uint rowIndex] => Cells.Where(x => x.Row == rowIndex).ToArray();

        public XlCell Get(uint col, uint row)
        {
            XlCell result = Cells.SingleOrDefault(x => x.Row == row && x.Col == col);
            if (result == null)
                throw new IndexOutOfRangeException(string.Format("Ячейка с индексом [{0},{1}] не найдена", row, col));
            else
                return result;
        }

        public XlCell Get(string Reference)
        {
            XlCell result = Cells.SingleOrDefault(x => x.Reference == Reference);
            if (result == null)
                throw new IndexOutOfRangeException(string.Format("Ячейка с индексом [{0}] не найдена", Reference));
            else
                return result;
        }

        public List<XlCell> Cells { get; set; }
        public uint Id { get; set; }
        //=================================================
        #endregion
    }
}