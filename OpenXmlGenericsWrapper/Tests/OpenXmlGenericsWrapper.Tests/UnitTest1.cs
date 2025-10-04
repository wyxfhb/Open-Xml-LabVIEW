using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlGenericsWrapper.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            Text txt = new DocumentFormat.OpenXml.Spreadsheet.Text("Otto");
            string result = txt.InnerText;
            Assert.AreEqual("Otto", result);
        }
    }
}
