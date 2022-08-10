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
        public FileController(GoogleDriveService googleDriveService)
        {
            this.googleDriveService = googleDriveService;
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
    }
}
