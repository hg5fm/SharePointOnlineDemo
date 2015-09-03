using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using File = Microsoft.SharePoint.Client.File;

namespace SharePointOnlineDemo
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            ConnectToDemoSite();
            UploadDemo();
            Console.ReadLine();
        }



        private static void ConnectToDemoSite()
        {
            const string siteUrl = "https://oauthplay-my.sharepoint.com/personal/allieb_oauthplay_onmicrosoft_com";
            using (var ctx = new ClientContext(siteUrl))
            {
                //Provide count and pwd for connecting to the source
                var passWord = new SecureString();
                foreach (char c in "Pastries101".ToCharArray()) passWord.AppendChar(c);
                ctx.Credentials = new SharePointOnlineCredentials("katiej@oauthplay.onmicrosoft.com", passWord);

                // Actual code for operations
                Web web = ctx.Web;
                ctx.Load(web);
                ctx.Load(ctx.Web.SiteUsers);
                ctx.ExecuteQuery();

                Console.WriteLine(string.Format("Connected to site with title of {0}", web.Title));
                Console.WriteLine("Site Users:");
                foreach (var siteUser in ctx.Web.SiteUsers)
                {
                    Console.WriteLine("{0} - {1} - {2}", siteUser.Id, siteUser.LoginName, siteUser.Title);
                }
                //Console.ReadLine();
            }
        }

        public static void UploadDemo()

        {
            const string siteUrl = "https://oauthplay.sharepoint.com";
            const string folderUrl = "/Shared Documents/DemoDocs";
            using (var ctx = new ClientContext(siteUrl))
            {
                var passWord = new SecureString();
                foreach (char c in "Pastries101".ToCharArray()) passWord.AppendChar(c);
                ctx.Credentials = new SharePointOnlineCredentials("katiej@oauthplay.onmicrosoft.com", passWord);


                var fileName = Path.GetTempFileName();
                System.IO.File.WriteAllText(fileName, "Test");
                var spFolder = ctx.Web.GetFolderByServerRelativeUrl(folderUrl);
                var file = UploadFileChunked(spFolder, fileName, "test.txt");
                var item = file.ListItemAllFields;


                const string template = "i:0#.f|membership|{0}@oauthplay.onmicrosoft.com";
                var modifiedBy = ctx.Web.EnsureUser(string.Format(template, "belindan"));
                var createdBy = ctx.Web.EnsureUser(string.Format(template, "allieb"));
                
                ctx.Load(modifiedBy);
                ctx.Load(createdBy);
                ctx.ExecuteQuery();

                item["Editor"] = modifiedBy.Id;
                var modifiedByField = new ListItemFormUpdateValue
                {
                    FieldName = "Modified_x0020_By",
                    FieldValue = modifiedBy.Id.ToString()
                };

                item["Author"] = createdBy.Id;
                var createdByField = new ListItemFormUpdateValue
                {
                    FieldName = "Created_x0020_By",
                    FieldValue = createdBy.Id.ToString()
                };

                item["Modified"] = DateTime.Now.Subtract(TimeSpan.FromDays(400)).ToUniversalTime();
                item["Created"] = DateTime.Now.Subtract(TimeSpan.FromDays(500)).ToUniversalTime();

                // it doesn't matter if you add both modifiedByField and createdByField.
                // As long as the list is non-empty all changes appear to carry over.
                var updatedValues = new List<ListItemFormUpdateValue> { modifiedByField, createdByField };
                item.ValidateUpdateListItem(updatedValues, true, string.Empty);
                ctx.Load(item);
                ctx.ExecuteQuery();


                System.IO.File.Delete(fileName);
                Console.WriteLine("Successfully uploaded document '{0}'", file.ServerRelativeUrl);
            }

        }


        public static File UploadFileChunked(Folder spFolder, string sourceFilePath, string uniqueFileName, int fileChunkSizeInMB = 3)
        {
            // Each sliced upload requires a unique ID.
            var uploadId = Guid.NewGuid();

            var ctx = (ClientContext)spFolder.Context;

            File uploadFile;

            var blockSize = fileChunkSizeInMB * 1024 * 1024;

            var fileSize = new FileInfo(sourceFilePath).Length;

            // Use regular approach.
            FileStream fs;
            if (fileSize <= blockSize)
            {
                using (fs = new FileStream(sourceFilePath, FileMode.Open))
                {
                    var fileInfo = new FileCreationInformation
                    {
                        ContentStream = fs,
                        Url = uniqueFileName,
                        Overwrite = true
                    };

                    uploadFile = spFolder.Files.Add(fileInfo);
                    ctx.Load(uploadFile);
                    ctx.ExecuteQuery();
                    // Return the file object for the uploaded file.
                    return uploadFile;
                }
            }
            // Use large file upload approach.

            fs = null;
            try
            {
                fs = System.IO.File.Open(sourceFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using (var br = new BinaryReader(fs))
                {
                    var buffer = new byte[blockSize];
                    byte[] lastBuffer = null;
                    long fileoffset = 0;
                    long totalBytesRead = 0;
                    int bytesRead;
                    var first = true;
                    var last = false;

                    // Read data from file system in blocks. 
                    while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        totalBytesRead = totalBytesRead + bytesRead;

                        // You've reached the end of the file.
                        if (totalBytesRead == fileSize)
                        {
                            last = true;
                            // Copy to a new buffer that has the correct size.
                            lastBuffer = new byte[bytesRead];
                            Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                        }

                        ClientResult<long> bytesUploaded = null;
                        if (first)
                        {
                            using (var contentStream = new MemoryStream())
                            {
                                // Add an empty file.
                                var fileInfo = new FileCreationInformation
                                {
                                    ContentStream = contentStream,
                                    Url = uniqueFileName,
                                    Overwrite = true
                                };
                                uploadFile = spFolder.Files.Add(fileInfo);

                                // Start upload by uploading the first slice. 
                                using (var s = new MemoryStream(buffer))
                                {
                                    // Call the start upload method on the first slice.
                                    bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                    ctx.ExecuteQuery();
                                    // fileoffset is the pointer where the next slice will be added.
                                    fileoffset = bytesUploaded.Value;
                                }

                                // You can only start the upload once.
                                first = false;
                            }
                        }
                        else
                        {
                            // Get a reference to your file.
                            uploadFile = ctx.Web.GetFileByServerRelativeUrl(spFolder.ServerRelativeUrl +
                                                                            Path.AltDirectorySeparatorChar +
                                                                            uniqueFileName);
                            if (last)
                            {
                                // Is this the last slice of data?
                                using (var s = new MemoryStream(lastBuffer))
                                {
                                    // End sliced upload by calling FinishUpload.
                                    uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                    ctx.ExecuteQuery();

                                    // Return the file object for the uploaded file.
                                    uploadFile = ctx.Web.GetFileByServerRelativeUrl(spFolder.ServerRelativeUrl +
                                                Path.AltDirectorySeparatorChar +
                                                uniqueFileName);
                                    return uploadFile;
                                }
                            }
                            else
                            {
                                using (var s = new MemoryStream(buffer))
                                {
                                    // Continue sliced upload.
                                    bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                    ctx.ExecuteQuery();
                                    // Update fileoffset for the next slice.
                                    fileoffset = bytesUploaded.Value;
                                }
                            }
                        }

                    } // while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                }
            }
            finally
            {
                if (fs != null)
                {
                    fs.Dispose();
                }
            }

            return null;
        }

    }


}
