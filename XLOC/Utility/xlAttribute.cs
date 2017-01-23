using System;
using System.Collections.Generic;
using XLOC.Utility.Extensions;

namespace XLOC.Utility
{
    [AttributeUsage(AttributeTargets.Property)]
    public class xlFieldAttribute : Attribute
    {
        #region Constructor
        //=================================================
        xlFieldAttribute(xlContentType contentType, bool isRequired)
        {
            ContentType = contentType;
            IsRequired = isRequired;
            Captions = new List<string>();
        }

        public xlFieldAttribute(xlContentType contentType, bool isRequired, params string[] captions) : this(contentType, isRequired) { captions.ForEach(Captions.Add); }
        public xlFieldAttribute(xlContentType contentType, bool isRequired, string caption) : this(contentType, isRequired, new string[] { caption }) { }
        public xlFieldAttribute(xlContentType contentType, params string[] captions) : this(contentType, true, captions) { }
        public xlFieldAttribute(bool isRequired, string caption) : this(xlContentType.SharedString, isRequired, caption) { }
        public xlFieldAttribute(bool isRequired, params string[] caption) : this(xlContentType.SharedString, isRequired, caption) { }
        public xlFieldAttribute(string caption) : this(true, caption) { }
        public xlFieldAttribute(params string[] caption) : this(true, caption) { }
        //=================================================
        #endregion

        public List<string> Captions { get; set; }
        public xlContentType ContentType { get; set; }
        public bool IsRequired { get; set; }
    }
}