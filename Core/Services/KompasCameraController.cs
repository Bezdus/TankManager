using KompasAPI7;
using System;
using System.Runtime.InteropServices;
using TankManager.Core.Constants;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Управляет камерой и позиционированием в KOMPAS
    /// </summary>
    class KompasCameraController
    {
        private readonly KompasContext _context;
        private readonly ILogger _logger;

        public KompasCameraController(KompasContext context, ILogger logger)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public void FocusOnPart(IPart7 part)
        {
            if (part == null)
                return;

            try
            {
                var gabarit = GetPartGabarit(part);
                var center = CalculateGlobalCenter(part);
                var scale = CalculateScale(gabarit);
                UpdateCamera(center, scale);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Failed to focus on part: {part.Name}", ex);
                throw;
            }
        }

        public (double x, double y, double z) CalculateGlobalCenter(IPart7 detail)
        {
            var gabarit = GetPartGabarit(detail);
            var localCenter = CalculateLocalCenter(gabarit);

            double globalX = localCenter.x;
            double globalY = localCenter.y;
            double globalZ = localCenter.z;

            IPart7 currentPart = TryGetParent(detail);

            while (currentPart != null)
            {
                IPart7 parentPart = TryGetParent(currentPart);

                if (parentPart != null)
                {
                    currentPart.Placement.GetOrigin(out double originX, out double originY, out double originZ);
                    globalX += originX;
                    globalY += originY;
                    globalZ += originZ;
                }

                currentPart = parentPart;
            }

            return (globalX, globalY, globalZ);
        }

        public static Gabarit GetPartGabarit(IPart7 part)
        {
            var gabarit = new Gabarit();
            part.GetGabarit(true, true,
                out gabarit.X1, out gabarit.Y1, out gabarit.Z1,
                out gabarit.X2, out gabarit.Y2, out gabarit.Z2);
            return gabarit;
        }

        private (double x, double y, double z) CalculateLocalCenter(Gabarit g)
        {
            return ((g.X1 + g.X2) / 2.0, (g.Y1 + g.Y2) / 2.0, (g.Z1 + g.Z2) / 2.0);
        }

        private double CalculateScale(Gabarit g)
        {
            double maxSize = Math.Max(
                Math.Max(Math.Abs(g.X2 - g.X1), Math.Abs(g.Y2 - g.Y1)),
                Math.Abs(g.Z2 - g.Z1));

            return maxSize > 0
                ? KompasConstants.ScaleFactor / maxSize
                : KompasConstants.DefaultScale;
        }

        private void UpdateCamera((double x, double y, double z) center, double scale)
        {
            object matrixObj = null;
            try
            {
                matrixObj = _context.ViewProjectionManager.Matrix3D;

                if (matrixObj is Array matrix)
                {
                    matrix.SetValue(center.x, KompasConstants.CameraMatrixXIndex);
                    matrix.SetValue(center.y, KompasConstants.CameraMatrixYIndex);
                    matrix.SetValue(center.z, KompasConstants.CameraMatrixZIndex);
                    _context.ViewProjectionManager.SetMatrix3D(matrix, scale);
                }
            }
            finally
            {
                if (matrixObj != null && Marshal.IsComObject(matrixObj))
                {
                    Marshal.ReleaseComObject(matrixObj);
                }
            }
        }

        private IPart7 TryGetParent(IPart7 part)
        {
            try
            {
                return part.Parent as IPart7;
            }
            catch
            {
                return null;
            }
        }

        internal struct Gabarit
        {
            public double X1, Y1, Z1, X2, Y2, Z2;

            public double SizeX => Math.Abs(X2 - X1);
            public double SizeY => Math.Abs(Y2 - Y1);
            public double SizeZ => Math.Abs(Z2 - Z1);

            public double MaxSize => Math.Max(SizeX, Math.Max(SizeY, SizeZ));
            public double Height => Math.Abs(Z2 - Z1);
        }
    }
}