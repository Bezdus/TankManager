using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TankManager.Core.Models;

namespace TankManager.Core.Services
{
    public interface IKompasService : IDisposable
    {
        List<PartModel> LoadDocument(string filePath);
        void ShowDetailInKompas(PartModel detail);
    }
}
