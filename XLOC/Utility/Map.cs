using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using XLOC.Utility.Extensions;
using System.Reflection;

namespace XLOC.Utility
{
    internal class Map<T> : Dictionary<string, MapItem>
    {
        #region Constructor
        //=================================================
        public Map(Dictionary<string, string> dictionary) : base()
        {
            foreach (var prop in getProperties())
            {
                try
                {
                    var attr = GetAttribute<xlFieldAttribute>(prop);
                    var tmp = dictionary.SingleOrDefault(x => x.Value.Equals(prop.Name, StringComparison.CurrentCultureIgnoreCase) || attr.Captions.Any(y => x.Value.Equals(y, StringComparison.CurrentCultureIgnoreCase)));
                    if (!isDefault(tmp))
                        this[tmp.Key.rMatch(@"^[A-Z]+")] = new MapItem { Attribute = attr, Property = prop };
                }
                catch (InvalidOperationException ex) { Exceptioins.Add(new InvalidOperationException($"Дублирование заголовков, {prop.Name}", ex)); }
            }

            var captions = GetClassCaptions(true);
            MissingFields = captions.Where(x => Values.All(y => !y.Contains(x))).Select(x => x.First());
        }
        //=================================================
        #endregion

        #region static
        //=================================================
        static string[][] GetClassCaptions(bool requiredOnly = false) => typeof(T).GetProperties()
            .Where(x => Attribute.IsDefined(x, typeof(xlFieldAttribute)) && !(requiredOnly && (isNullable(x) || !isRequired(x))))
            .Select(x => GetAttribute<xlFieldAttribute>(x).Captions.Union(new string[] { x.Name }).ToArray()).ToArray();
        static IEnumerable<PropertyInfo> getProperties() => typeof(T).GetProperties().Where(x => Attribute.IsDefined(x, typeof(xlFieldAttribute)));
        static IEnumerable<Cell> GetCaptionCells(WorksheetPart sheet) => sheet.Worksheet.GetFirstChild<SheetData>().Descendants<Row>().First().Descendants<Cell>().Where(x => x.CellValue != null);
        static AtttrType GetAttribute<AtttrType>(PropertyInfo x) where AtttrType : Attribute => (AtttrType)Attribute.GetCustomAttribute(x, typeof(AtttrType));
        static bool isRequired(PropertyInfo x) => GetAttribute<xlFieldAttribute>(x).IsRequired;
        static bool isNullable(PropertyInfo x) => x.PropertyType.IsGenericType && x.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>);
        static bool isDefault<type>(type item) => item.Equals(default(type));
        //=================================================
        #endregion

        #region Methods
        //=================================================
        public IEnumerable<Cell> GetCells(Row row) => row.Descendants<Cell>().Where(x => Keys.Any(y => x.CellReference.Value.rIsMatch(y)));
        //=================================================
        #endregion

        #region Properties
        //=================================================
        public IEnumerable<string> MissingFields { get; set; }
        public List<Exception> Exceptioins { get; set; } = new List<Exception>();
        public bool IsValid => !MissingFields.Any() || Exceptioins.Any();
        //=================================================
        #endregion
    }
}