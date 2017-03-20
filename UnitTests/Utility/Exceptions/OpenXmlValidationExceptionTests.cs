using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;

namespace XLOC.Utility.Exceptions.Tests
{
    [TestClass]
    public class OpenXmlValidationExceptionTests
    {
        [TestMethod]
        public void OpenXmlValidationExceptionTest()
        {
            var res = new OpenXmlValidationException(new[] { new DocumentFormat.OpenXml.Validation.ValidationErrorInfo() }, "a");
            Assert.AreEqual("a", res.Message);
            Assert.AreEqual(1, res.ValidationErrors.Count());
            Assert.AreEqual(1, res.ValidationErrors.Count());
        }
    }
}