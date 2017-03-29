using XLOC.Utility.Extensions;
using System;
using System.Collections.Generic;

namespace XLOC.Book
{
    public class XlBook
    {
        #region Variables
        //=================================================
        string _name = string.Empty;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        public XlBook() => Sheets = new List<XlSheet>();
        //=================================================
        #endregion

        #region Methods
        //=================================================
        public void SaveAs(string path) => XlConverter.FromBook(this).SaveToFile(path);

        public XlSheet AddSheet(string name)
        {
            var result = new XlSheet((uint)Sheets.Count + 1, name);
            Sheets.Add(result);
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
                if (value.rIsMatch(@"[\\\/\:\*\?\""\<\>\|]"))
                    throw new ArgumentException("Имя книги содержит недопустимые символы");
                _name = string.Format("{0}.xlsx", value.rMatch(@"\w+(?=\.\w+$)?"));
            }
        }
        public List<XlSheet> Sheets { get; set; }
        public XlSheet this[int index] => Sheets[index];
        //=================================================
        #endregion
    }
}