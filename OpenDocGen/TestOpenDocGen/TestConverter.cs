using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using OpenDocGen;

namespace TestOpenDocGen
{
    [TestClass]
    public class TestConverter
    {

        private TestContext testContextInstance;
        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }


        [TestMethod]
        public void TestConvert2Html()
        {
            /**
             * This example loads each document into a byte array, then into a memory stream, so that the document can be opened for writing without
             * modifying the source document.
             */
            var docPath = "../../data";
            var outPath = "../../out";
            Assert.IsTrue(Directory.Exists(docPath));

            foreach (var file in Directory.GetFiles(docPath, "*.docx"))
            {
                TestContext.WriteLine($"Converting Document {file} to Path: {outPath}");
                var converter = new HtmlConverter();
                converter.ConvertToHtml(file, outPath);
            }
        }
    }
}
