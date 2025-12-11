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
    }
}