using Kompas6Constants;
using Kompas6Constants3D;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using TankManager.Core.Constants;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Определяет пересечения объектов между камерой и целевой деталью
    /// </summary>
    internal class PartIntersectionDetector
    {
        private const double CameraDistance = 5000.0; // Расстояние от камеры до центра (мм)
        private const int RaycastSteps = 100;
        private const double SearchEpsilon = 50.0;
        private const double MaxDistanceMultiplier = 0.99;
        private const double TargetDimensionMultiplier = 2.0;

        public static ILineSegment3D testLine;

        private readonly KompasContext _context;
        private readonly ILogger _logger;
        private readonly KompasCameraController _cameraController;

        public PartIntersectionDetector(KompasContext context, ILogger logger, KompasCameraController cameraController)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _cameraController = cameraController ?? throw new ArgumentNullException(nameof(cameraController));
        }

        public List<IPart7> FindIntersectingObjects(IPart7 targetDetail)
        {
            if (targetDetail == null || _context.Document == null)
                return new List<IPart7>();

            var center = _cameraController.CalculateGlobalCenter(targetDetail);
            var objects = new List<object>();
            var uniqueParts = new HashSet<IntPtr>();

            object matrixObj = null;
            IFindObject3DParameters filterParam = null;

            try
            {
                matrixObj = _context.ViewProjectionManager.Matrix3D;

                if (matrixObj is Array matrix && matrix.Length >= 16)
                {
                    var viewDirection = ExtractViewDirection(matrix);
                    var cameraPosition = CalculateCameraPosition(center, viewDirection);
                    var rayDirection = CalculateRayDirection(center, cameraPosition, out double distance);

                    // Создаем тестовую линию для визуализации
                    //CreateTestLine(cameraPosition, center);

                    filterParam = CreateFindObjectParameters();
                    var filter = filterParam as FindObject3DParameters;

                    IPart7 topPart = _context.Document.TopPart as IPart7;
                    IntPtr targetPtr = Marshal.GetIUnknownForObject(targetDetail);

                    var targetGabarit = _cameraController.GetPartGabarit(targetDetail);
                    double targetMaxDim = CalculateMaxDimension(targetGabarit);

                    _logger.LogInfo($"Camera position: ({cameraPosition.x:F2}, {cameraPosition.y:F2}, {cameraPosition.z:F2})");
                    _logger.LogInfo($"Target center: ({center.x:F2}, {center.y:F2}, {center.z:F2})");
                    _logger.LogInfo($"Ray direction: ({rayDirection.x:F3}, {rayDirection.y:F3}, {rayDirection.z:F3})");
                    _logger.LogInfo($"Distance: {distance:F2}, MaxDim: {targetMaxDim:F2}");

                    // ИСПОЛЬЗУЕМ ГЛОБАЛЬНЫЕ КООРДИНАТЫ - как в старой функции
                    PerformRaycast(topPart, cameraPosition, rayDirection, distance,
                        filter, targetPtr, targetDetail, center, targetMaxDim,
                        objects, uniqueParts);

                    Marshal.Release(targetPtr);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError("Error finding intersecting objects", ex);
            }
            finally
            {
                CleanupResources(matrixObj, filterParam, uniqueParts);
            }


            List<IPart7> intersectingParts = new List<IPart7>();

            foreach (var obj in objects)
            {
                if (obj is IModelObject part)
                {
                    if (part.ModelObjectType == ksObj3dTypeEnum.o3d_face)
                    {
                        intersectingParts.Add(part.Part);
                    }
                }
            }

            _logger.LogInfo($"Found {objects.Count} intersecting objects");
            return intersectingParts;
        }

        private void CreateTestLine((double x, double y, double z) cameraPosition, (double x, double y, double z) center)
        {
            try
            {
                IAuxiliaryGeomContainer auxiliaryGeomContainer = _context.TopPart as IAuxiliaryGeomContainer;
                if (auxiliaryGeomContainer != null)
                {
                    testLine = auxiliaryGeomContainer.LineSegments3D.Add();
                    testLine.BuildingType = ksLineSegment3DTypeEnum.ksLSTTwoPoints;
                    testLine.SetPoint(true, cameraPosition.x, cameraPosition.y, cameraPosition.z);
                    testLine.SetPoint(false, center.x, center.y, center.z);
                    testLine.Update();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Failed to create test line: {ex.Message}");
            }
        }

        private (double x, double y, double z) ExtractViewDirection(Array matrix)
        {
            double viewDirX = Convert.ToDouble(matrix.GetValue(8));
            double viewDirY = Convert.ToDouble(matrix.GetValue(9));
            double viewDirZ = Convert.ToDouble(matrix.GetValue(10));

            double length = Math.Sqrt(viewDirX * viewDirX + viewDirY * viewDirY + viewDirZ * viewDirZ);

            return (viewDirX / length, viewDirY / length, viewDirZ / length);
        }

        private (double x, double y, double z) CalculateCameraPosition(
            (double x, double y, double z) center,
            (double x, double y, double z) viewDirection)
        {
            return (
                center.x + viewDirection.x * CameraDistance,
                center.y + viewDirection.y * CameraDistance,
                center.z + viewDirection.z * CameraDistance
            );
        }

        private (double x, double y, double z) CalculateRayDirection(
            (double x, double y, double z) center,
            (double x, double y, double z) cameraPosition,
            out double distance)
        {
            double dirX = center.x - cameraPosition.x;
            double dirY = center.y - cameraPosition.y;
            double dirZ = center.z - cameraPosition.z;

            distance = Math.Sqrt(dirX * dirX + dirY * dirY + dirZ * dirZ);

            return (dirX / distance, dirY / distance, dirZ / distance);
        }

        private IFindObject3DParameters CreateFindObjectParameters()
        {
            IKompasDocument1 doc1 = _context.Document as IKompasDocument1;
            var filterParam = doc1.GetInterface(KompasAPIObjectTypeEnum.ksObjectFindObject3DParameters)
                as IFindObject3DParameters;
            filterParam.ModelObjectType = ksObj3dTypeEnum.o3d_face;

            return filterParam;
        }

        private double CalculateMaxDimension(KompasCameraController.Gabarit gabarit)
        {
            return Math.Max(
                Math.Max(Math.Abs(gabarit.X2 - gabarit.X1),
                         Math.Abs(gabarit.Y2 - gabarit.Y1)),
                Math.Abs(gabarit.Z2 - gabarit.Z1));
        }

        private void PerformRaycast(
            IPart7 topPart,
            (double x, double y, double z) cameraPosition,
            (double x, double y, double z) rayDirection,
            double totalDistance,
            FindObject3DParameters filter,
            IntPtr targetPtr,
            IPart7 targetDetail,
            (double x, double y, double z) targetCenter,
            double targetMaxDim,
            List<object> objects,
            HashSet<IntPtr> uniqueParts)
        {
            int foundCount = 0;
            
            for (int i = 0; i < RaycastSteps; i++)
            {
                double t = (totalDistance / RaycastSteps) * i;

                // Проверяем, находится ли точка между камерой и центром детали
                if (t > totalDistance * MaxDistanceMultiplier) 
                    continue;

                double pointX = cameraPosition.x + rayDirection.x * t;
                double pointY = cameraPosition.y + rayDirection.y * t;
                double pointZ = cameraPosition.z + rayDirection.z * t;

                int beforeCount = objects.Count;
                ProcessRaycastPoint(topPart, pointX, pointY, pointZ, filter,
                    targetPtr, targetDetail, targetCenter, targetMaxDim,
                    totalDistance, objects, uniqueParts);

                if (objects.Count > beforeCount)
                {
                    foundCount++;
                    _logger.LogInfo($"Step {i}/{RaycastSteps}: Found object at ({pointX:F2}, {pointY:F2}, {pointZ:F2}), t={t:F2}");
                }
            }
            
            _logger.LogInfo($"Total raycast steps: {RaycastSteps}, hits: {foundCount}");
        }

        private void ProcessRaycastPoint(
            IPart7 topPart,
            double pointX, double pointY, double pointZ,
            FindObject3DParameters filter,
            IntPtr targetPtr,
            IPart7 targetDetail,
            (double x, double y, double z) targetCenter,
            double targetMaxDim,
            double totalDistance,
            List<object> objects,
            HashSet<IntPtr> uniqueParts)
        {
            object result = null;
            try
            {
                result = topPart.FindObjectsByPointWithParam(
                    pointX, pointY, pointZ,
                    false,
                    SearchEpsilon,
                    filter
                );

                if (result is Array objectsArray && objectsArray.Length > 0)
                {
                    _logger.LogInfo($"FindObjectsByPoint returned {objectsArray.Length} objects at ({pointX:F2}, {pointY:F2}, {pointZ:F2})");
                    
                    ProcessFoundObjects(objectsArray, targetPtr, targetDetail,
                        pointX, pointY, pointZ, targetCenter, targetMaxDim,
                        totalDistance, objects, uniqueParts);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Error in FindObjectsByPointWithParam: {ex.Message}");
            }
            finally
            {
                if (result != null && Marshal.IsComObject(result))
                {
                    Marshal.ReleaseComObject(result);
                }
            }
        }

        private void ProcessFoundObjects(
            Array objectsArray,
            IntPtr targetPtr,
            IPart7 targetDetail,
            double pointX, double pointY, double pointZ,
            (double x, double y, double z) targetCenter,
            double targetMaxDim,
            double totalDistance,
            List<object> objects,
            HashSet<IntPtr> uniqueParts)
        {
            foreach (var obj in objectsArray)
            {
                if (obj == null) continue;

                if (obj is IModelObject modelObj &&
                    modelObj.ModelObjectType == ksObj3dTypeEnum.o3d_face)
                {
                    ProcessFaceObject(modelObj, targetPtr, targetDetail,
                        pointX, pointY, pointZ, targetCenter, targetMaxDim,
                        totalDistance, objects, uniqueParts);
                }
            }
        }

        private void ProcessFaceObject(
            IModelObject modelObj,
            IntPtr targetPtr,
            IPart7 targetDetail,
            double pointX, double pointY, double pointZ,
            (double x, double y, double z) targetCenter,
            double targetMaxDim,
            double totalDistance,
            List<object> objects,
            HashSet<IntPtr> uniqueParts)
        {
            IPart7 objPart = modelObj.Part;
            if (objPart == null) return;

            IntPtr partPtr = Marshal.GetIUnknownForObject(objPart);

            try
            {
                // Проверяем, что это не целевая деталь и не дубликат
                if (partPtr != targetPtr && !uniqueParts.Contains(partPtr))
                {
                    // Проверяем, что это не родитель целевой детали
                    if (!IsParentOf(objPart, targetDetail))
                    {
                        // Вычисляем расстояние от найденной точки до центра целевой детали
                        double distToTarget = CalculateDistance(
                            pointX, pointY, pointZ,
                            targetCenter.x, targetCenter.y, targetCenter.z);

                        // Добавляем только если объект находится между камерой и целью
                        // и расстояние меньше размера целевой детали
                        if (distToTarget < totalDistance && 
                            distToTarget < targetMaxDim * TargetDimensionMultiplier)
                        {
                            uniqueParts.Add(partPtr);
                            objects.Add(modelObj);
                            _logger.LogInfo($"Added intersecting part: {objPart.Name}, distToTarget={distToTarget:F2}");
                            return; // Не освобождаем partPtr, он сохранён в HashSet
                        }
                        else
                        {
                            Marshal.Release(partPtr);
                        }
                    }
                    else
                    {
                        Marshal.Release(partPtr);
                    }
                }
                else
                {
                    if (partPtr != targetPtr)
                        Marshal.Release(partPtr);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Error processing face object: {ex.Message}");
                if (!uniqueParts.Contains(partPtr) && partPtr != targetPtr)
                {
                    Marshal.Release(partPtr);
                }
            }
        }

        private double CalculateDistance(double x1, double y1, double z1, double x2, double y2, double z2)
        {
            return Math.Sqrt(
                Math.Pow(x1 - x2, 2) +
                Math.Pow(y1 - y2, 2) +
                Math.Pow(z1 - z2, 2)
            );
        }

        private bool IsParentOf(IPart7 potentialParent, IPart7 child)
        {
            IPart7 current = TryGetParent(child);
            IntPtr parentPtr = Marshal.GetIUnknownForObject(potentialParent);

            try
            {
                while (current != null)
                {
                    IntPtr currentPtr = Marshal.GetIUnknownForObject(current);
                    bool isSame = (currentPtr == parentPtr);
                    Marshal.Release(currentPtr);

                    if (isSame)
                        return true;

                    current = TryGetParent(current);
                }
            }
            finally
            {
                Marshal.Release(parentPtr);
            }

            return false;
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

        private void CleanupResources(object matrixObj, IFindObject3DParameters filterParam, HashSet<IntPtr> uniqueParts)
        {
            if (matrixObj != null && Marshal.IsComObject(matrixObj))
            {
                Marshal.ReleaseComObject(matrixObj);
            }

            if (filterParam != null && Marshal.IsComObject(filterParam))
            {
                Marshal.ReleaseComObject(filterParam);
            }

            foreach (var ptr in uniqueParts)
            {
                Marshal.Release(ptr);
            }
        }
    }
}