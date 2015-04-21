using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using System.IO;

using Microsoft.SharePoint.Client;

namespace SharePointHelper
{
    public class SharePointHelper
    {
        private ClientContext ctx;
        private Web WebRoot;

        public string SPUrl { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path">SharePoint website path</param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        public SharePointHelper(string path, string username, string password)
        {
            SPUrl = path;
            Username = username;
            Password = password;
            ctx = Authenticate(path, username, password);
            WebRoot = ctx.Web;

            ctx.Load(WebRoot,
                     w => w.Title,
                     w => w.ServerRelativeUrl);
            ctx.ExecuteQuery();
        }

        /// <summary>
        /// Authenticate client context
        /// </summary>
        /// <param name="path">SharePoint website path</param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        /// <returns></returns>
        public ClientContext Authenticate(string path, string username, string password)
        {
            ClientContext clientCtx = new ClientContext(path);
            NetworkCredential credentials = new NetworkCredential(username, password);
            clientCtx.Credentials = credentials;
            return clientCtx;
        }

        /// <summary>
        /// Retrieves site by title
        /// </summary>
        /// <param name="siteTitle"></param>
        /// <returns></returns>
        public Web GetWebByTitle(string siteTitle)
        {
            try
            {
                var query = ctx.LoadQuery(WebRoot.Webs.Where(p => p.Title == siteTitle));
                ctx.ExecuteQuery();
                return query.FirstOrDefault();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Retrieves document library
        /// </summary>
        /// <param name="siteTitle"></param>
        /// <param name="libraryName"></param>
        /// <returns></returns>
        public List GetDocumentLibrary(string siteTitle, string libraryName)
        {
            List items = null;

            try
            {
                var web = GetWebByTitle(siteTitle);
                if (web != null)
                {
                    items = web.Lists.GetByTitle(libraryName);
                    ctx.Load(items, w => w.Id, w => w.Title, w => w.Fields, w => w.RootFolder);
                    ctx.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return items;
        }

        public List<Microsoft.SharePoint.Client.File> GetAllFiles(string siteTitle, string libraryName)
        {
            try
            {
                var list = GetDocumentLibrary(siteTitle, libraryName);
                var files = list.RootFolder.Files;
                ctx.Load(files);
                ctx.ExecuteQuery();
                List<Microsoft.SharePoint.Client.File> fileList = files.ToList();
                return fileList;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void CreateFolder(string siteTitle, string libraryName, string folder)
        {
            try
            {
                var list = GetDocumentLibrary(siteTitle, libraryName);
                if (list != null)
                {
                    var folders = list.RootFolder.Folders;
                    ctx.Load(folders);
                    ctx.ExecuteQuery();
                    var newFolder = folders.Add(folder);
                    ctx.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public ListItemCollection GetDocumentByName(string documentName, string libraryName)
        {
            try
            {
                List SPLibraryList = ctx.Web.Lists.GetByTitle(libraryName);
                ListItemCollection listItems = SPLibraryList.GetItems(QueryFileName(documentName, false));
                ctx.Load(SPLibraryList);
                ctx.Load(listItems);
                ctx.ExecuteQuery();
                return listItems;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public MemoryStream DownloadFileDirect(string SharepointFilePath)
        {
            try
            {
                WebClient client = new WebClient();
                client.Credentials = new NetworkCredential(Username, Password);
                MemoryStream FileMemoryStream = new MemoryStream(client.DownloadData(SharepointFilePath));
                client.Dispose();
                return FileMemoryStream;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void UploadDocument(string siteTitle, string libraryName, string fileLocPath)
        {
            try
            {
                string targetLoc = string.Empty;

                FileInfo fInfo = new FileInfo(fileLocPath);

                if (siteTitle != "/")
                {
                    Web site = GetWebByTitle(siteTitle);
                    targetLoc = string.Format("{0}/{1}/{2}", site.ServerRelativeUrl, libraryName, fInfo.Name);
                }
                else
                {
                    targetLoc = string.Format("{0}/{1}/{2}", "", libraryName, fInfo.Name);
                }

                using (var fs = new FileStream(fileLocPath, FileMode.Open))
                {
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, targetLoc, fs, true);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public ListItemCollection GetDocumentByMetadata(string siteTitle, string libraryName, string metadataName, string metadataValue, bool subFolders)
        {
            return QueryLibrary(siteTitle, libraryName, QueryEqual(metadataName.Replace(" ", "_x0020_"), metadataValue, "Text"));
        }

        private ListItemCollection QueryLibrary(string siteTitle, string libraryName, CamlQuery libraryQuery)
        {
            try
            {
                List SPLibraryList = GetDocumentLibrary(siteTitle, libraryName);
                ListItemCollection listItems = SPLibraryList.GetItems(libraryQuery);
                ctx.Load(listItems);
                ctx.ExecuteQuery();

                return listItems;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static CamlQuery QueryEqual(string field, string value, string valueType)
        {
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml +=
                                              @"<View>
                                                 <Query>
                                                  <Where>
                                                    <Eq>
                                                      <FieldRef Name='" + field + "'/>" +
                                                  "   <Value Type='" + valueType + "'>" + value + "</Value>" +
                                                 "  </Eq>" +
                                               "  </Where>" +
                                              " </Query>" +
                                            " </View>";
            return camlQuery;
        }

        private static CamlQuery QueryFileName(string fileName, bool subFolders)
        {
            CamlQuery camlQuery = new CamlQuery();
            if (subFolders)
                camlQuery.ViewXml = @"<View Scope='Recursive'>";
            else
                camlQuery.ViewXml = @"<View>";
            camlQuery.ViewXml =
                                         @"<Query>
                                          <Where>
                                            <Eq>
                                              <FieldRef Name='FileLeafRef'/>
                                              <Value Type='Text'>" + fileName + "</Value>" +
                                            "</Eq>" +
                                         "</Where>" +
                                        "</Query>" +
                                     "</View>";
            return camlQuery;
        }
    }
}