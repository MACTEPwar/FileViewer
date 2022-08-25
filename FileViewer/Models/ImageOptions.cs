using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FileViewer.Models
{
    public class CanvasOptions
    {
        // Длинна холста в см
        public int Width { get; set; } = 3;
        // Высота холста в см
        public int Height { get; set; } = 3;

        public List<ImageOptions> ImageOptions { get; set; } = new List<ImageOptions>();

    }

    public class ImageOptions
    {
        public int Width { get; set; } = 3;
        // Высота холста в см
        public int Height { get; set; } = 3;
        public string ImagePath { get; set; }
    }

}
