using qXlsxLib.Utility;
using qXlsxLib.Utility.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;

namespace qXlsxLib.Book
{
    public class xlSheet
    {
        #region Variables
        //=================================================
        string _name = string.Empty;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        xlSheet(uint Id)
        {
            Cells = new List<xlCell>();
            this.Id = Id;
        }

        internal xlSheet(uint Id, string ListName)
            : this(Id)
        {
            Name = ListName;
        }
        //=================================================
        #endregion

        #region Methods
        //=================================================
        public xlCell AddCell()
        {
            var result = new xlCell();
            Cells.Add(result);
            return result;
        }

        public xlCell AddCell(object cellValue, int aRow, int aCol, xlContentType aType)
        {
            var result = AddCell();
            result.Value = cellValue;
            result.Row = aRow;
            result.Col = aCol;
            result.Type = aType;
            return result;
        }

        public xlCell AddCell(object cellValue, string aReference, xlContentType aType)
        {
            var result = AddCell();
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
            get { return _name; }
            set
            {
                if (!value.rIsMatch(@"^[^\:\\\/\?\*\[\]]+$"))
                    throw new ArgumentException("Имя книги содержит недопустимые символы");
                _name = value;
            }
        }

        public IEnumerable<xlCell> this[uint rowIndex]
        {
            get
            {
                return Cells.Where(x => x.Row == rowIndex).ToArray();
            }
        }

        public xlCell Get(uint col, uint row)
        {
            xlCell result = Cells.SingleOrDefault(x => x.Row == row && x.Col == col);
            if (result == null)
                throw new IndexOutOfRangeException(string.Format("Ячейка с индексом [{0},{1}] не найдена", row, col));
            else
                return result;
        }

        public xlCell Get(string Reference)
        {
                xlCell result = Cells.SingleOrDefault(x => x.Reference == Reference);
                if (result == null)
                    throw new IndexOutOfRangeException(string.Format("Ячейка с индексом [{0}] не найдена", Reference));
                else
                    return result;
        }

        public List<xlCell> Cells { get; set; }
        public uint Id { get; set; }
        //=================================================
        #endregion
    }
}