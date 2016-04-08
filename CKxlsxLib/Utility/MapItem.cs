using qXlsxLib.Utility;
using System.Collections.Generic;

namespace qXlsxLib.Utility
{
    internal class MapItem
    {
        public System.Reflection.PropertyInfo Property { get; set; }
        public xlFieldAttribute Attribute { get; set; }
    }
}