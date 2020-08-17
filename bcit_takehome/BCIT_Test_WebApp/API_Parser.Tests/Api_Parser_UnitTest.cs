using System.Web;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using API_Parser.Controllers;
using System.Net.Http;
using System.Net;
using System.Web.Http;

namespace API_Parser.Tests
{
    [TestClass]
    public class Api_Parser_UnitTest
    {
        ParserController pc;
        Api_Parser_UnitTest()
        {
            pc = new ParserController();
            pc.Request = new HttpRequestMessage();
            pc.Request.SetConfiguration(new HttpConfiguration());
        }
        [TestMethod]
        public void EmptyFilePathTest()
        {            
            HttpResponseMessage response = pc.Get("");
            Assert.AreEqual(HttpStatusCode.NotFound, response.StatusCode);
        }

        [TestMethod]
        public void IncorrectFilePathTest()
        {            
            HttpResponseMessage response = pc.Get(@"C:\temp.docx");
            Assert.AreEqual(HttpStatusCode.NotFound, response.StatusCode);
        }
    }
}
