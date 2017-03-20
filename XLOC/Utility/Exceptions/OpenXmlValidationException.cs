using DocumentFormat.OpenXml.Validation;
using System;
using System.Runtime.Serialization;

namespace XLOC.Utility.Exceptions
{
    [Serializable]
    public class OpenXmlValidationException : Exception
    {
        public OpenXmlValidationException(ValidationErrorInfo[] errorsInfo, string message) : base(message) => ValidationErrors = errorsInfo;
        public ValidationErrorInfo[] ValidationErrors { get; set; }
        public override void GetObjectData(SerializationInfo info, StreamingContext context) => base.GetObjectData(info, context);
    }
}