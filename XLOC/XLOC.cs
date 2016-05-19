using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XLOC.Book;
using XLOC.Reader;

namespace XLOC
{
    public class XlConverter
    {
        public static XLOCReader FromStream(Stream stream, XLOCConfiguration configuration = null) => new XLOCReader { Configuration = configuration ?? new XLOCConfiguration(), Document = SpreadsheetDocument.Open(stream, false) };

        public static XLOCReader FromFile(string path, XLOCConfiguration configuration = null)
        {
            try { return FromStream(new MemoryStream(File.ReadAllBytes(path)), configuration); }
            catch (Exception ex) { throw new IOException(string.Format("Не удалось открыть файл {0}", path), ex); }
        }
    }

    public class XLOCReader
    {
        #region Properties
        //=================================================
        public XLOCConfiguration Configuration { get; set; }
        public SpreadsheetDocument Document { get; internal set; }
        //=================================================
        #endregion

        #region Methods
        //=================================================
        public IEnumerable<T> ReadToEnumerable<T>() where T : Utility.IxlCompatible, new() => new xlArrayReader(Configuration).ReadToEnumerable<T>(Document);
        public T[] ReadToArray<T>() where T : Utility.IxlCompatible, new() => ReadToEnumerable<T>().ToArray();
        public xlBook ReadToBook() => new xlBookReader(Document).ReadToBook();
        //=================================================
        #endregion
    }
}