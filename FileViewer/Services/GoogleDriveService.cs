using FileViewer.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Drive.v2.Data;
using Google.Apis.Services;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FileViewer.Services
{
    public class GoogleDriveService
    {
        private DriveService driveService = null;

        public GoogleDriveService(GoogleOAuthService googleOAuthService)
        {
            driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = googleOAuthService.GetCurrentCredential(),
                ApplicationName = "My Project 53610",
            });
        }

        public FileList GetFilesByFolder(string folderId = "root")
        {
            folderId = folderId == "root" ? "0AKi2z3Ku7rbnUk9PVA" : folderId;
            var listQuery = driveService.Files.List();
            //listQuery.Fields = "items(id, title, mimeType, parents)";
            listQuery.Fields = "*";
            listQuery.Q = $"'{folderId}' in parents and title != '$_TEMPLATES_$'";
            return listQuery.Execute();
        }

        //public FileList GetPathMetadata(List<string> ids)
        //{
        //    var listQuery = driveService.Files.List();
        //    listQuery.Fields = "items(id, title, mimeType, parents)";
        //    //listQuery.Q = $"'{ids[0]}' in id";
        //    //listQuery.Q = $"id = '{ids[0]}'";
        //    listQuery.Q = $"id = '{ids[0]}'";
        //    return listQuery.Execute();
        //}

        public IList<Google.Apis.Drive.v2.Data.File> GetFiles()
        {
            FilesResource.ListRequest listRequest = driveService.Files.List();
            listRequest.Fields = "*";
            IList<Google.Apis.Drive.v2.Data.File> files = listRequest.Execute()
                   .Items;
            return files;
        }

        // 1NCcAefsdpmaVh8mV5NzxVFJFevyjckc2P
        public void ChangePermission(string fileId)
        {
            Google.Apis.Drive.v2.Data.Permission permission = new Google.Apis.Drive.v2.Data.Permission()
            {
                Role = "writer",
                Type = "anyone",
                WithLink = true
            };

            driveService.Permissions.Insert(permission, fileId).Execute();
        }

        public void Share(string fileId)
        {
            Google.Apis.Drive.v2.Data.File body = new Google.Apis.Drive.v2.Data.File();
            body.Shared = true;
            driveService.Files.Patch(body, fileId).Execute();
        }

        public Google.Apis.Drive.v2.Data.File UploadFile(IFormFile file, string folderId = null)
        {
            if (file.Length > 0)
            {
                Google.Apis.Drive.v2.Data.Permission permission = new Google.Apis.Drive.v2.Data.Permission()
                {
                    Role = "writer",
                    Type = "anyone",
                    WithLink = true
                };

                Google.Apis.Drive.v2.Data.File body = new Google.Apis.Drive.v2.Data.File();
                body.Title = file.FileName;
                body.Description = "Uploaded with .NET!";
                body.Shared = true;
                body.MimeType = "application/vnd.google-apps.document";
                body.Permissions = new List<Google.Apis.Drive.v2.Data.Permission>() { permission };
                body.Parents = new List<ParentReference>()
                {
                    new ParentReference()
                    {
                        Id = folderId ??  "0AKi2z3Ku7rbnUk9PVA",
                        IsRoot = folderId == null
                    }
                };
                // body.Parents = new List<string> { _parent };// UN comment if you want to upload to a folder(ID of parent folder need to be send as paramter in above method)
                using (var stream = file.OpenReadStream())
                {
                    try
                    {
                        FilesResource.InsertMediaUpload request = driveService.Files.Insert(body, stream, "text/plain");
                        request.SupportsTeamDrives = true;
                        // You can bind event handler with progress changed event and response recieved(completed event)
                        //request.ProgressChanged += Request_ProgressChanged;
                        //request.ResponseReceived += Request_ResponseReceived;
                        request.Upload();
                        return request.ResponseBody;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error Occured");
                        return null;
                    }
                }

            }
            else
            {
                Console.WriteLine("The file does not exist.");
                return null;
            }
        }

        public string CreateFolder(Folder folderMetadata)
        {
            var fileMetadata = new Google.Apis.Drive.v2.Data.File()
            {
                Title = folderMetadata.Title,
                MimeType = "application/vnd.google-apps.folder",
                Parents = new List<ParentReference>() {
                    new ParentReference()
                    {
                        Id = folderMetadata.ParentId ?? "0AKi2z3Ku7rbnUk9PVA",
                        IsRoot = folderMetadata.ParentId == null
                    }
                }
            };

            var request = driveService.Files.Insert(fileMetadata);
            request.Fields = "id";
            var file = request.Execute();
            return file.Id;
        }
    }
}
