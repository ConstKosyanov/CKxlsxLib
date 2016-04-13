using XLOC.Utility;
using System.Collections.Generic;
using System.Linq;
using System;

namespace XLOC.Utility
{
    internal class MapItem
    {
        static EQ eq = new EQ();
        class EQ : EqualityComparer<string>
        {
            public override bool Equals(string x, string y) => string.Equals(x, y, StringComparison.CurrentCultureIgnoreCase);
            public override int GetHashCode(string obj) => obj.Length;
        }
        internal bool Contains(IEnumerable<string> items) => Attribute.Captions.Intersect(items, eq).Any() || items.Contains(Property.Name, eq);

        public System.Reflection.PropertyInfo Property { get; set; }
        public xlFieldAttribute Attribute { get; set; }
    }
}