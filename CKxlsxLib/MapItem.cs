using System.Collections.Generic;

namespace CKxlsxLib
{
    internal class MapItem
    {
        public System.Reflection.PropertyInfo Property { get; set; }
        public xlFieldAttribute Attribute { get; set; }
    }
}