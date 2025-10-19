using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace TankManager.Core.Services
{
    public class ComObjectManager : IDisposable
    {
        private readonly List<object> _comObjects = new List<object>();
        private readonly ILogger _logger;

        public ComObjectManager(ILogger logger)
        {
            _logger = logger;
        }

        public T Track<T>(T comObject) where T : class
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                _comObjects.Add(comObject);
            }
            return comObject;
        }

        public void Release(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
            {
                try
                {
                    _comObjects.Remove(comObject);
                    Marshal.ReleaseComObject(comObject);
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"Failed to release COM object: {ex.Message}");
                }
            }
        }

        public void Dispose()
        {
            foreach (var comObject in _comObjects)
            {
                try
                {
                    if (comObject != null && Marshal.IsComObject(comObject))
                    {
                        Marshal.ReleaseComObject(comObject);
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning($"Failed to dispose COM object: {ex.Message}");
                }
            }
            _comObjects.Clear();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}