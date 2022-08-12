using FileViewer.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FileViewer.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        readonly GoogleDriveService googleDriveService = null;
        readonly GoogleDocsService googleDocsService = null;
        readonly WordInterop wordInteropService = null;
        public FileController(GoogleDriveService googleDriveService, GoogleDocsService googleDocsService, WordInterop wordInteropService)
        {
            this.googleDriveService = googleDriveService;
            this.googleDocsService = googleDocsService;
            this.wordInteropService = wordInteropService;
        }

        [HttpPost("upload"), DisableRequestSizeLimit]
        public async Task<IActionResult> FileUpload([FromQuery] IFormFile file, [FromQuery] string parentFolderId)
        {

            return new JsonResult(this.googleDriveService.UploadFile(file, parentFolderId));
        }

        [HttpDelete("delete/{id}")]
        public async Task<IActionResult> DeleteFile(string id)
        {
            return BadRequest("No implement");
        }

        [HttpGet("test/{id}")]
        public async Task<IActionResult> Test(string id)
        {
            try
            {
                this.googleDocsService.test(id);
                return Ok();
            }
            catch(Exception e)
            {
                return BadRequest("Bad");
            }
        }

        [HttpGet("test2/{id}")]
        public async Task<IActionResult> Test2(string id)
        {
            try
            {
                //this.googleDocsService.test2(id);
                this.googleDocsService.insertImage(id);
                return Ok();
            }
            catch (Exception e)
            {
                return BadRequest("Bad");
            }
        }

        [HttpGet("AddPageNumbers")]
        public async Task<IActionResult> AddPageNumbers()
        {
            try
            {
                //this.googleDocsService.test2(id);
                this.wordInteropService.OpenDocument();
                return Ok();
            }
            catch (Exception e)
            {
                return BadRequest("Bad");
            }
        }

        [HttpGet("test3")]
        public async Task<IActionResult> test3()
        {
            try
            {
                //this.googleDocsService.test2(id);
                this.wordInteropService.Test();
                return Ok();
            }
            catch (Exception e)
            {
                return BadRequest("Bad");
            }
        }
    }
}
