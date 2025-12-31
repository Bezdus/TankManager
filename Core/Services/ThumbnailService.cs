using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media.Imaging;
using Microsoft.WindowsAPICodePack.Shell;


namespace TankManager.Core.Services
{
    public class ThumbnailService
    {
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes,
            ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }

        private const uint SHGFI_ICON = 0x100;
        private const uint SHGFI_LARGEICON = 0x0;

        public static BitmapSource GetFileThumbnail(string filePath, int width = 512, int height = 512)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                return null;

            try
            {
                using (ShellFile shellFile = ShellFile.FromFilePath(filePath))
                {
                    Bitmap thumbnail = shellFile.Thumbnail.ExtraLargeBitmap ??
                                      shellFile.Thumbnail.LargeBitmap ??
                                      shellFile.Thumbnail.MediumBitmap;

                    if (thumbnail != null)
                    {
                        var bitmapSource = ConvertBitmapToBitmapSource(thumbnail, width, height);
                        thumbnail.Dispose(); // Освобождаем ресурсы
                        return bitmapSource;
                    }
                }
            }
            catch
            {
                // Fallback к иконке файла
                return GetFileIcon(filePath);
            }

            return GetFileIcon(filePath);
        }

        public static BitmapSource GetFileIcon(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return null;

            SHFILEINFO shinfo = new SHFILEINFO();
            IntPtr hImgSmall = SHGetFileInfo(filePath, 0, ref shinfo,
                (uint)Marshal.SizeOf(shinfo), SHGFI_ICON | SHGFI_LARGEICON);

            if (shinfo.hIcon != IntPtr.Zero)
            {
                try
                {
                    Icon icon = Icon.FromHandle(shinfo.hIcon);
                    BitmapSource bitmapSource = Imaging.CreateBitmapSourceFromHIcon(
                        icon.Handle,
                        Int32Rect.Empty,
                        BitmapSizeOptions.FromEmptyOptions());

                    bitmapSource.Freeze();
                    return bitmapSource;
                }
                finally
                {
                    DestroyIcon(shinfo.hIcon);
                }
            }

            return null;
        }

        private static BitmapSource ConvertBitmapToBitmapSource(Bitmap bitmap, int width, int height)
        {
            IntPtr hBitmap = bitmap.GetHbitmap();
            try
            {
                BitmapSource bitmapSource = Imaging.CreateBitmapSourceFromHBitmap(
                    hBitmap,
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromWidthAndHeight(width, height));

                bitmapSource.Freeze();
                return bitmapSource;
            }
            finally
            {
                DeleteObject(hBitmap);
            }
        }
    }
}
