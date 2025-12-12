namespace TankManager.Core.Models
{
    /// <summary>
    /// Тип продукта/детали
    /// </summary>
    public enum ProductType
    {
        /// <summary>
        /// Обычная деталь
        /// </summary>
        Part,

        /// <summary>
        /// Покупная деталь
        /// </summary>
        PurchasedPart,

        /// <summary>
        /// Листовой прокат
        /// </summary>
        SheetMaterial,

        /// <summary>
        /// Трубный прокат
        /// </summary>
        TubularProduct
    }
}