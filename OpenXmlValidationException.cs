using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CKxlsxLib
{
    public class OpenXmlValidationException : Exception
    {
        public OpenXmlValidationException(ValidationErrorInfo[] errorsInfo, string message) : base(message) { ValidationErrors = errorsInfo; }

        ValidationErrorInfo[] ValidationErrors { get; set; }
    }
}