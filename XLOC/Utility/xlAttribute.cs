using System;
using System.Collections.Generic;
using XLOC.Utility.Extensions;

namespace XLOC.Utility
{
    [AttributeUsage(AttributeTargets.Property)]
    public class XlFieldAttribute : Attribute
    {
        #region Constructor
        //=================================================
        XlFieldAttribute(XlContentType contentType, bool isRequired)
        {
            ContentType = contentType;
            IsRequired = isRequired;
            Captions = new List<string>();
        }

        public XlFieldAttribute(XlContentType contentType, bool isRequired, params string[] captions) : this(contentType, isRequired) => captions.ForEach(Captions.Add);
        public XlFieldAttribute(XlContentType contentType, bool isRequired, string caption) : this(contentType, isRequired, new[] { caption }) { }
        public XlFieldAttribute(XlContentType contentType, params string[] captions) : this(contentType, true, captions) { }
        public XlFieldAttribute(bool isRequired, string caption) : this(XlContentType.SharedString, isRequired, caption) { }
        public XlFieldAttribute(bool isRequired, params string[] caption) : this(XlContentType.SharedString, isRequired, caption) { }
        public XlFieldAttribute(params string[] caption) : this(true, caption) { }
        //=================================================
        #endregion

        public List<string> Captions { get; set; }
        public XlContentType ContentType { get; set; }
        public bool IsRequired { get; set; }
    }
}