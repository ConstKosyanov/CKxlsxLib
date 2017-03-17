using XLOC.Utility;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace XLOC.Book
{
    public class XlCell
    {
        #region Variables
        //=================================================
        int row, column;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        public XlCell() { }
        public XlCell(string reference, XlContentType? type, object value, int? sharedId)
        {
            Reference = reference;
            Type = type;
            Value = value;
            SharedId = sharedId;
        }
        //=================================================
        #endregion

        #region Private
        //=================================================
        internal static string GetReference(int col, int row)
        {
            int m;
            string name = string.Empty;
            while (col > 0)
            {
                m = (col - 1) % 26;
                name = (char)(65 + m) + name;
                col = (int)((col - m) / 26);
            }
            return name + row;
        }
        //=================================================
        #endregion

        #region Properties
        //=================================================
        public object Value { get; set; }
        public int Row
        {
            get => row;
            set
            {
                if (value < 1)
                    throw new ArgumentOutOfRangeException("Индекс строки должен быть больше 1");
                row = value;
            }
        }
        public int Col
        {
            get => column;
            set
            {
                if (value < 1)
                    throw new ArgumentOutOfRangeException("Индекс столбца должен быть больше 1");
                column = value;
            }
        }
        public string Reference
        {
            get => GetReference(column, row);
            set
            {
                if (!Regex.IsMatch(value, @"^[A-Z]+\d+$"))
                    throw new ArgumentException("Адрес ячейки должен соответсвовать формату [A-Z]+[0-9]+");

                int i = Regex.Match(value, "^[A-Z]+").Value.Count() - 1;
                int col = 0;
                foreach (var item in Regex.Match(value, "^[A-Z]+").Value.Select(x => (byte)x - 64))
                {
                    col += (int)(item * Math.Pow(26, i--));
                }
                Col = col;
                Row = int.Parse(Regex.Match(value, @"\d+$").Value);
            }
        }
        public XlContentType? Type { get; set; }
        public int? SharedId { get; internal set; }
        //=================================================
        #endregion
    }
}