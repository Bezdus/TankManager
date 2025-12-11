using System;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    public interface IKompasService : IDisposable
    {
        Product LoadDocument(string filePath);
        Product LoadActiveDocument();
        void ShowDetailInKompas(PartModel detail);
    }
}
