using System.Collections.ObjectModel;
using System.Linq;
using KompasAPI7;
using TankManager.Core.Services;

namespace TankManager.Core.Models
{
    /// <summary>
    /// Представляет изделие (сборку) с его составными частями
    /// </summary>
    public class Product : PartModel
    {
        public Product() : base()
        {
            Details = new ObservableCollection<PartModel>();
            Materials = new ObservableCollection<MaterialInfo>();
            StandardParts = new ObservableCollection<PartModel>();

        }

        public Product(IPart7 part, KompasContext context, int instanceIndex = 0) 
            : base(part, context, instanceIndex)
        {
            Details = new ObservableCollection<PartModel>();
            Materials = new ObservableCollection<MaterialInfo>();
            StandardParts = new ObservableCollection<PartModel>();

        }

        public ObservableCollection<PartModel> Details { get; }
        public ObservableCollection<MaterialInfo> Materials { get; }
        public ObservableCollection<PartModel> StandardParts { get; }


        public void Clear()
        {
            Details.Clear();
            Materials.Clear();
            StandardParts.Clear();
            Name = null;
            Marking = null;
            Mass = 0;
        }
    }
}