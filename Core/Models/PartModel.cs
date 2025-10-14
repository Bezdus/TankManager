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

        public PartModel(IPart7 part, KompasContext context)
        {
            Part = part;
            Name = part.Name;
            Marking = part.Marking;
            DetailType = GetDetailType(part, context?.SpecificationSectionProperty);
            Material = part.Material;
            Mass = Part.Mass/1000;
            FilePath = part.FileName;

            LoadDocumentPreview();
        }

        public PartModel(IBody7 body, KompasContext context)
        {
            Name = body.Name;
        }

        private string GetDetailType(IPart7 part, IProperty property)
        {
            if (property == null)
                return null;

            object markingObj;
            bool fromSource;
            IPropertyKeeper propertyKeeper = (IPropertyKeeper)part;
            propertyKeeper.GetPropertyValue((KompasAPI7._Property)property, out markingObj, false, out fromSource);
            return markingObj?.ToString();
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
