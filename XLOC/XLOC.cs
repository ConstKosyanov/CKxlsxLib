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
        public static XLOCReader FromStream(Stream stream, XLOCConfiguration configuration = null) => new XLOCReader((configuration ?? new XLOCConfiguration()).AddDocument(SpreadsheetDocument.Open(stream, false)));

        public static XLOCReader FromFile(string path, XLOCConfiguration configuration = null)
        {
            try { return FromStream(new MemoryStream(File.ReadAllBytes(path)), configuration); }
            catch (Exception ex) { throw new IOException(string.Format("Не удалось открыть файл {0}", path), ex); }
        }

        public static XLOCReader FromBuffer(byte[] buf, XLOCConfiguration configuration = null) => FromStream(new MemoryStream(buf), configuration);
    }

    public class XLOCReader
    {
        #region Constructor
        //=================================================
        internal XLOCReader(XLOCConfiguration Configuration)
        {
            this.Configuration = Configuration;
        }
        //=================================================
        #endregion

        #region Properties
        //=================================================
        public XLOCConfiguration Configuration { get; set; }
        //=================================================
        #endregion

        #region Methods
        //=================================================
        public IEnumerable<T> ReadToEnumerable<T>() where T : new() => new xlArrayReader(Configuration).ReadToEnumerable<T>();
        public T[] ReadToArray<T>() where T : new() => ReadToEnumerable<T>().ToArray();
        public xlBook ReadToBook() => new xlBookReader(Configuration).ReadToBook(Configuration.Document);
        //=================================================
        #endregion
    }
}