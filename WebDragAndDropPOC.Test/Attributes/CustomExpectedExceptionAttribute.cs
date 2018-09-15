using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace CustomAttributes
{
    public sealed class CustomExpectedExceptionAttribute : ExpectedExceptionBaseAttribute
    {

        public Type ExpectedExceptionType { get; private set; }
        public string ExpectedExceptionMessage { get; set; }
        public string ParameterName { get; set; }

        public CustomExpectedExceptionAttribute(Type expectedType)
        {
            ExpectedExceptionType = expectedType;
        }

        public CustomExpectedExceptionAttribute(Type expectedType, string expectedMessage)
        {
            ExpectedExceptionType = expectedType;
            ExpectedExceptionMessage = expectedMessage;
        }

        public CustomExpectedExceptionAttribute(Type expectedType, string expectedMessage, string parameterName)
        {
            ExpectedExceptionType = expectedType;
            ExpectedExceptionMessage = expectedMessage;
            ParameterName = parameterName;
        }

        protected override void Verify(Exception exception)
        {
            Assert.IsNotNull(exception);
            Assert.IsInstanceOfType(exception, ExpectedExceptionType);

            if (!string.IsNullOrWhiteSpace(ExpectedExceptionMessage))
            {
                Assert.AreEqual(ExpectedExceptionMessage, exception.Message);
            }

            if (exception is ArgumentException argumentException && !string.IsNullOrWhiteSpace(ParameterName))
            {
                Assert.AreEqual(ParameterName, argumentException.ParamName);
            }
        }
    }
}
