using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Положение и размеры главного окна, ширины колонок
    /// </summary>
    [DataContract]
    public class WindowSettings
    {
        [DataMember] public double Left { get; set; }
        [DataMember] public double Top { get; set; }
        [DataMember] public double Width { get; set; }
        [DataMember] public double Height { get; set; }
        [DataMember] public bool IsMaximized { get; set; }

        /// <summary>
        /// Пропорции трёх колонок основного контента (звёздочные ширины)
        /// </summary>
        [DataMember] public double[] ColumnStars { get; set; }
    }

    /// <summary>
    /// Загрузка и сохранение настроек окна в window_settings.json рядом с exe
    /// </summary>
    public class WindowSettingsService
    {
        private static readonly string SettingsFilePath =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "window_settings.json");

        private readonly ILogger _logger;

        public WindowSettingsService(ILogger logger)
        {
            _logger = logger;
        }

        public WindowSettings Load()
        {
            try
            {
                if (!File.Exists(SettingsFilePath))
                    return null;

                var serializer = new DataContractJsonSerializer(typeof(WindowSettings));
                using (var memoryStream = new MemoryStream(File.ReadAllBytes(SettingsFilePath)))
                {
                    return (WindowSettings)serializer.ReadObject(memoryStream);
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning($"Ошибка загрузки настроек окна: {ex.Message}");
                return null;
            }
        }

        public void Save(WindowSettings settings)
        {
            try
            {
                var serializer = new DataContractJsonSerializer(typeof(WindowSettings));
                using (var memoryStream = new MemoryStream())
                {
                    serializer.WriteObject(memoryStream, settings);
                    AtomicFile.WriteAllBytes(SettingsFilePath, memoryStream.ToArray());
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning($"Ошибка сохранения настроек окна: {ex.Message}");
            }
        }
    }
}
