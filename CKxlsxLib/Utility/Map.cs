using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using qXlsxLib.Utility.Extensions;
using System.Reflection;

namespace qXlsxLib.Utility
{
    internal class Map<T> : Dictionary<string, MapItem>
    {
        public Map(Dictionary<string, string> dictionary) : base()
        {
            foreach (var p in getProperties())
            {
                var attr = GetAttribute<xlFieldAttribute>(p);
                var tmp = dictionary.SingleOrDefault(x => x.Value.Equals(p.Name, StringComparison.CurrentCultureIgnoreCase) || attr.Captions.Any(y => x.Value.Equals(y, StringComparison.CurrentCultureIgnoreCase)));
                if (tmp != default(typeof(KeyValuePair<string,string>)))
                {

                }
            }
        }

        static IEnumerable<PropertyInfo> getProperties() => typeof(T).GetProperties().Where(x => Attribute.IsDefined(x, typeof(xlFieldAttribute)));
        static IEnumerable<Cell> GetCaptionCells(WorksheetPart sheet) => sheet.Worksheet.GetFirstChild<SheetData>().Descendants<Row>().First().Descendants<Cell>().Where(x => x.CellValue != null);
        static AtttrType GetAttribute<AtttrType>(PropertyInfo x) where AtttrType : Attribute => (AtttrType)Attribute.GetCustomAttribute(x, typeof(AtttrType));
        public IEnumerable<Cell> GetCells(Row row) => row.Descendants<Cell>().Where(x => Keys.Any(y => x.CellReference.Value.rIsMatch(y)));
    }
}