using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using TankManager.Core.Services;

namespace TankManager.Core.Models
{
    public class PartModel
    {
        public IPart7 Part { get; set; }
        public string Name { get; set; }
        public string Marking { get; set; }
        public string DetailType { get; set; }
        public string Material { get; set; }
        public double Mass { get; set; }

        public string FilePath { get; set; }

        public BitmapSource FilePreview { get; set; }

        public PartModel(IPart7 part, string detailType)
        {
            Part = part;
            Name = part.Name;
            Marking = part.Marking;
            DetailType = detailType;
            Material = part.Material;
            Mass = Part.Mass/1000;
            FilePath = part.FileName;

            LoadDocumentPreview();
        }

        private void LoadDocumentPreview()
        {
            if (string.IsNullOrEmpty(FilePath) || !System.IO.File.Exists(FilePath))
            {
                FilePreview = null;
                return;
            }

            try
            {
                // Получаем иконку/превью файла
                FilePreview = ThumbnailService.GetFileThumbnail(FilePath);
            }
            catch (Exception ex)
            {
                FilePreview = null;
            }
        }
    }
}
