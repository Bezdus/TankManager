namespace TankManager.Core.Constants
{
    public static class KompasConstants
    {
        // Типы деталей
        public const string DetailsSectionName = "Детали";
        public const string StandardPartsType = "Стандартные изделия";
        public const string OtherPartsType = "Прочие изделия";
        public const string PurchasedPartType = "Покупная деталь";
        public const string PartType = "Деталь";

        // Свойства
        public const string MaterialPropertyName = "Материал";
        public const string MassPropertyName = "Масса";
        public const string SpecificationSectionPropertyName = "Раздел спецификации";

        // Масштаб и камера
        public const double DefaultScale = 1.0;
        public const double ScaleFactor = 100.0;
        public const double MassConversionFactor = 1000.0;

        // Индексы матрицы камеры
        public const int CameraMatrixXIndex = 12;
        public const int CameraMatrixYIndex = 13;
        public const int CameraMatrixZIndex = 14;

        // Расширения файлов
        public const string KompasFileExtension = ".a3d";
    }
}