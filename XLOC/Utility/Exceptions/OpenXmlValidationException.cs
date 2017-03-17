using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;

namespace XLOC.Utility.Exceptions
{
    [Serializable]
    public class OpenXmlValidationException : Exception
    {
        public OpenXmlValidationException(ValidationErrorInfo[] errorsInfo, string message) : base(message) => ValidationErrors = errorsInfo;
        ValidationErrorInfo[] ValidationErrors { get; set; }

        public override void GetObjectData(SerializationInfo info, StreamingContext context) => base.GetObjectData(info, context);
    }
}