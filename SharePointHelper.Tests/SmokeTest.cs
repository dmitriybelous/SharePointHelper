using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.SharePoint.Client;

namespace SharePointHelper.Tests
{
    [TestClass]
    public class SmokeTest
    {
        string path = "";
        string username = "";
        string password = "";
        string site = "";
        string library = "";

        string document = "test.doc";
        string metadataName = "Year";
        string metadataValue = "2015";

        SharePointHelper spHelper = null;

        [ClassInitialize()]
        public static void ClassInit(TestContext context)
        {
            
        }

        public SmokeTest()
        {
            spHelper = new SharePointHelper(path, username, password);
        }

        [TestMethod]
        public void Authenticate()
        {
            Assert.AreNotEqual(null, spHelper.Authenticate(path, username, password));
        }

        [TestMethod]
        public void GetWebByTitle()
        {
            Assert.AreNotEqual(null, spHelper.GetWebByTitle(site));
        }

        [TestMethod]
        public void GetDocumentLibrary()
        {
            Assert.AreNotEqual(null, spHelper.GetDocumentLibrary(site, library));
        }

        //[TestMethod]
        //public void GetAllFiles()
        //{
        //    List<File> files = spHelper.GetAllFiles(site, library);
        //    Assert.AreNotEqual(0, files.Count());
        //}

        [TestMethod]
        public void GetDocumentByName()
        {
            ListItemCollection files = spHelper.GetDocumentByName(document, library);
            Assert.AreNotEqual(0, files.ToList().Count);
        }

        [TestMethod]
        public void GetDocumentByMetadata()
        {
            ListItemCollection files = spHelper.GetDocumentByMetadata(site, library, metadataName, metadataValue, true);
            Assert.AreNotEqual(0, files.ToList().Count);
        }

        [ClassCleanup()]
        public static void ClassCleanup()
        {
           
        }
    }
}
