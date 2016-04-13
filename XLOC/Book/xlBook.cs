using XLOC.Utility.Extensions;
using System;
using System.Collections.Generic;

namespace XLOC.Book
{
    public class xlBook
    {
        #region Variables
        //=================================================
        string _name = string.Empty;
        //=================================================
        #endregion

        #region Constructor
        //=================================================
        public xlBook()
        {
            Sheets = new List<xlSheet>();
        }
        //=================================================
        #endregion

        #region Methods
        //=================================================
        public void SaveAs(string path)
        {
            Writer.xlWriter.Create(this).SaveToFile(path);
        }

        public xlSheet AddSheet(string name)
        {
            var result = new xlSheet((uint)Sheets.Count + 1, name);
            Sheets.Add(result);
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
                if (value.rIsMatch(@"[\\\/\:\*\?\""\<\>\|]"))
                    throw new ArgumentException("Имя книги содержит недопустимые символы");
                _name = string.Format("{0}.xlsx", value.rMatch(@"\w+(?=\.\w+$)?"));
            }
        }
        public List<xlSheet> Sheets { get; set; }
        public xlSheet this[int index] { get { return Sheets[index]; } }
        //=================================================
        #endregion
    }
}