using FileViewer.Models;
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
    public class FolderController : ControllerBase
    {
        readonly GoogleDriveService googleDriveService = null;
        public FolderController(GoogleDriveService googleDriveService)
        {
            this.googleDriveService = googleDriveService;
        }

        [HttpGet("{id=root}")]
        public async Task<IActionResult> GetChilderFilesByFolder(string id = "root")
        {
            return new JsonResult(this.googleDriveService.GetFilesByFolder(id));
        }

        //[HttpGet("{id}")]
        //[HttpGet("test")]
        //public async Task<IActionResult> GetPathMetadata()
        //{
        //    string id = "1t7QqgCYT_fPWueo66HWVWWr8J_0VwLou";
        //    return new JsonResult(this.googleDriveService.GetPathMetadata(new List<string>() { id }));
        //}

        [HttpPost]
        public async Task<IActionResult> CreateFolder([FromBody] Folder folderMetadata)
        {
            return new JsonResult(this.googleDriveService.CreateFolder(folderMetadata));
        }

        [HttpPut("{id}")]
        public async Task<IActionResult> EditFolder(string id, Google.Apis.Drive.v2.Data.File folderMetadata)
        {
            return BadRequest("No implement");
        }

        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteFolder(string id)
        {
            return BadRequest("No implement");
        }
    }
}
