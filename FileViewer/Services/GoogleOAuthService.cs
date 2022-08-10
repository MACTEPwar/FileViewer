using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.WindowsAzure.Storage.Shared.Protocol;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace FileViewer.Services
{
    public class GoogleOAuthService
    {
        private const string ClientId = "738387477576-k80tvb20ts5d3pbnr0poe710ssj85gev.apps.googleusercontent.com";
        private const string ClientSecret = "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";

        private ServiceAccountCredential credential = null;


        public GoogleOAuthService()
        {
            SignIn().Wait();
        }

        public async Task SignIn()
        {
            using (var certificate = new X509Certificate2("./key.p12", "notasecret", X509KeyStorageFlags.Exportable))
            {
                credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer("test-362@complete-silo-233622.iam.gserviceaccount.com")
                {
                    Scopes = new[] {
                        @"https://www.googleapis.com/auth/devstorage.read_write",
                        DriveService.Scope.Drive,
                        "https://www.googleapis.com/auth/drive"
                    }
                }.FromCertificate(certificate));
            }
        }

        public ServiceAccountCredential GetCurrentCredential() {
            return this.credential;
        }

        public static async Task Login()
        {
            using (var certificate = new X509Certificate2("./key.p12", "notasecret", X509KeyStorageFlags.Exportable))
            {
                var credential = new ServiceAccountCredential(new ServiceAccountCredential.Initializer("test-362@complete-silo-233622.iam.gserviceaccount.com")
                {
                    Scopes = new[] {
                        @"https://www.googleapis.com/auth/devstorage.read_write",
                        DriveService.Scope.Drive,
                        "https://www.googleapis.com/auth/drive"
                    }
                }.FromCertificate(certificate));

                /* Use Google Service with this authentication*/
                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "My Project 53610",
                });

                //GoogleOAuthService.uploadFile(service, "./file3.txt");
                // 17UywQ2kZlbpp1b08xitIw1wj_g3-0KC6

                GoogleOAuthService.Share(service, "17UywQ2kZlbpp1b08xitIw1wj_g3-0KC6");
                GoogleOAuthService.AddPermission(service, "17UywQ2kZlbpp1b08xitIw1wj_g3-0KC6");

                FilesResource.ListRequest listRequest = service.Files.List();
                //listRequest.PageSize = 10;
                //listRequest.Fields = "nextPageToken, files(id, name)";
                listRequest.Fields = "*";

                IList<Google.Apis.Drive.v2.Data.File> files = listRequest.Execute()
                    .Items;

                Console.WriteLine("OK");
            }
        }

        public static void Share(DriveService service, string fileId)
        {
            Google.Apis.Drive.v2.Data.File body = new Google.Apis.Drive.v2.Data.File();
            body.Title = System.IO.Path.GetFileName("File3");
            body.Description = "Uploaded with .NET!";
            body.Shared = true;
            body.MimeType = "text/plain";

            service.Files.Patch(body, fileId).Execute();
        }

        public static void AddPermission(DriveService service, string fileId)
        {
            Google.Apis.Drive.v2.Data.Permission permission = new Google.Apis.Drive.v2.Data.Permission()
            {
                Role = "owner",
                Type = "anyone",
                WithLink = true
            };

            //service.Permissions.Insert(permission, "17UywQ2kZlbpp1b08xitIw1wj_g3-0KC6").Execute();

            permission = new Google.Apis.Drive.v2.Data.Permission()
            {
                Role = "organizer",
                Type = "anyone",
                WithLink = true
            };

            //service.Permissions.Insert(permission, "17UywQ2kZlbpp1b08xitIw1wj_g3-0KC6").Execute();

            permission = new Google.Apis.Drive.v2.Data.Permission()
            {
                Role = "fileOrganizer",
                Type = "anyone",
                WithLink = true
            };

            //service.Permissions.Insert(permission, "17UywQ2kZlbpp1b08xitIw1wj_g3-0KC6").Execute();

            permission = new Google.Apis.Drive.v2.Data.Permission()
            {
                Role = "writer",
                Type = "anyone",
                WithLink = true
            };

            service.Permissions.Insert(permission, "17UywQ2kZlbpp1b08xitIw1wj_g3-0KC6").Execute();

            permission = new Google.Apis.Drive.v2.Data.Permission()
            {
                Role = "reader",
                Type = "anyone",
                WithLink = true
            };


            service.Permissions.Insert(permission, "17UywQ2kZlbpp1b08xitIw1wj_g3-0KC6").Execute();
        }


        public static Google.Apis.Drive.v2.Data.File uploadFile(DriveService _service, string _uploadFile, string _parent = null, string _descrp = "Uploaded with .NET!")
        {
            if (System.IO.File.Exists(_uploadFile))
            {
                Google.Apis.Drive.v2.Data.Permission permission = new Google.Apis.Drive.v2.Data.Permission()
                {
                    Role = "reader",
                    Type = "anyone",
                    WithLink = true
                };

                Google.Apis.Drive.v2.Data.File body = new Google.Apis.Drive.v2.Data.File();
                body.Title = System.IO.Path.GetFileName(_uploadFile);
                body.Description = _descrp;
                body.Shared = true;
                body.MimeType = "text/plain";
                body.Permissions = new List<Google.Apis.Drive.v2.Data.Permission>() { permission };
                // body.Parents = new List<string> { _parent };// UN comment if you want to upload to a folder(ID of parent folder need to be send as paramter in above method)
                byte[] byteArray = System.IO.File.ReadAllBytes(_uploadFile);
                System.IO.MemoryStream stream = new System.IO.MemoryStream(byteArray);
                try
                {
                    //_service.Files.
                    FilesResource.InsertMediaUpload request = _service.Files.Insert(body, stream, "text/plain");
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
            else
            {
                Console.WriteLine("The file does not exist.");
                return null;
            }
        }


        public static string GeneateOAuthRequestGenerate(string scope, string redirectUrl, string codeChallange)
        {
            var oAuthServerEndpoint = "https://accounts.google.com/o/oauth2/v2/auth";

            var queryParams = new Dictionary<string, string>
            {
                {"client_id", ClientId },
                {"redirect_uri", redirectUrl },
                {"response_type", "code" },
                {"scope", scope },
                {"code_challenge", codeChallange },
                {"code_challenge_method", "S256" },
                {"access_type","offline" }
            };

            var url = QueryHelpers.AddQueryString(oAuthServerEndpoint, queryParams);
            return url;
        }

        public static async Task<TokenResult> ExchangeCodeOnTokenAsync(string code, string codeVerifier, string redirectUrl)
        {
            var oAuthServerEndpoint = "https://oauth2.googleapis.com/token";

            var queryParams = new Dictionary<string, string>
            {
                {"client_id", ClientId },
                {"client_secret", ClientSecret },
                {"code", code },
                {"code_verifier", codeVerifier },
                {"grant_type", "authorization_code" },
                {"redirect_uri", redirectUrl }
            };


            var tokenResult = await QueryService.PostRequestAsync<TokenResult>(oAuthServerEndpoint, queryParams);
            return tokenResult;
        }

        public static async Task<TokenResult> RefreshTokenAsync(string refreshToken)
        {
            var refreshEndpoint = "https://oauth2.googleapis.com/token";

            var refreshParams = new Dictionary<string, string>
            {
                {"client_id", ClientId },
                {"client_secret", ClientSecret },
                {"grant_type", "refresh_token" },
                {"refresh_token", refreshToken }
            };

            var tokenResult = await QueryService.PostRequestAsync<TokenResult>(refreshEndpoint, refreshParams);
            return tokenResult;
        }
    }
}
