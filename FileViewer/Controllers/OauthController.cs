using FileViewer.Helpers;
using FileViewer.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FileViewer.Controllers
{
    [Route("[controller]")]
    [ApiController]
    public class OauthController : ControllerBase
    {
        readonly GoogleOAuthService googleOAuthService = null;
        readonly GoogleDriveService googleDriveService = null;
        public OauthController(GoogleOAuthService googleOAuthService, GoogleDriveService googleDriveService)
        {
            this.googleOAuthService = googleOAuthService;
            this.googleDriveService = googleDriveService;
        }

        [HttpGet("test")]
        public async Task<IActionResult> Test()
        {
            //await GoogleOAuthService.Login();
            return Content("test");
        }

        [HttpGet("test2")]
        public async Task<IActionResult> Test2()
        {
            return new JsonResult(this.googleDriveService.GetFiles());
        }

        [HttpGet("changePerm")]
        public async Task<IActionResult> ChangePerm()
        {
            this.googleDriveService.ChangePermission("1fIrVDmqz_3nJIRikd-YD4vTDwp6a_FlmDNQNqqpTu6A");
            this.googleDriveService.Share("1fIrVDmqz_3nJIRikd-YD4vTDwp6a_FlmDNQNqqpTu6A");
            return Ok();
        }

        //[HttpGet("uploadFile")]
        //public async Task<IActionResult> UploadFile()
        //{

        //    return new JsonResult(this.googleDriveService.UploadFile("./file4.txt"));
        //}

        //[HttpPost("response")]
        //public async Task<IActionResult> Post([FromBody] object obj)
        //{
        //    return new JsonResult(obj);
        //}
        [HttpPost("RedirectInOAuthServer")]

        public IActionResult RedirectInOAuthServer()
        {
            var scope = "https://www.googleapis.com/auth/drive";
            var redirectUrl = "http://localhost:20657/Oauth/CodeAsync";

            var codeVerifier = Guid.NewGuid().ToString();

            HttpContext.Session.SetString("codeVerifier", codeVerifier);

            var codeChallenge = Sha256Helper.ComputeHash(codeVerifier);

            var url = GoogleOAuthService.GeneateOAuthRequestGenerate(scope, redirectUrl, codeChallenge);
            return new JsonResult(new
            {
                url = url
            });
        }

        [HttpGet("CodeAsync")]
        public async Task<IActionResult> CodeAsync(string code)
        {
            string codeVerifier = HttpContext.Session.GetString("codeVerifier");
            string redirectUrl = "http://localhost:20657/Oauth/CodeAsync";

            var tokenResult = await GoogleOAuthService.ExchangeCodeOnTokenAsync(code, codeVerifier, redirectUrl);

            var refreshedTokenResult = await GoogleOAuthService.RefreshTokenAsync(tokenResult.RefreshToken);

            return Redirect("localhost://4200");
        }
    }
}
