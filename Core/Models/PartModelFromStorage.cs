using System.ComponentModel;

namespace TankManager.Core.Models
{
    /// <summary>
    /// PartModel восстановленный из хранилища
    /// </summary>
    public class PartModelFromStorage : PartModel
    {
        public PartModelFromStorage() : base()
        {
        }

        // Открываем сеттеры для восстановления из DTO
        public new string DetailType
        {
            get => base.DetailType;
            set => base.DetailType = value;
        }

        public new string PartId
        {
            get => base.PartId;
            set => base.PartId = value;
        }

        public new bool IsBodyBased
        {
            get => base.IsBodyBased;
            set => base.IsBodyBased = value;
        }

        public new int InstanceIndex
        {
            get => base.InstanceIndex;
            set => base.InstanceIndex = value;
        }

        public new double MetalCost
        {
            get => base.MetalCost;
            set => base.MetalCost = value;
        }

        public new double OperationsCost
        {
            get => base.OperationsCost;
            set => base.OperationsCost = value;
        }

        public new double TotalCost
        {
            get => base.TotalCost;
            set => base.TotalCost = value;
        }
    }
}