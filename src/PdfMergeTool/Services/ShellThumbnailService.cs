using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace PdfMergeTool.Services;

public static class ShellThumbnailService
{
    public static ImageSource? TryLoadThumbnail(string path, int width, int height)
    {
        try
        {
            var itemId = typeof(IShellItemImageFactory).GUID;
            SHCreateItemFromParsingName(path, IntPtr.Zero, ref itemId, out var factory);
            var size = new NativeSize { Width = width, Height = height };
            factory.GetImage(size, ShellItemImageFactoryOptions.ThumbnailOnly | ShellItemImageFactoryOptions.BiggerSizeOk, out var bitmapHandle);

            if (bitmapHandle == IntPtr.Zero)
            {
                return null;
            }

            try
            {
                var source = Imaging.CreateBitmapSourceFromHBitmap(
                    bitmapHandle,
                    IntPtr.Zero,
                    Int32Rect.Empty,
                    BitmapSizeOptions.FromEmptyOptions());
                source.Freeze();
                return source;
            }
            finally
            {
                DeleteObject(bitmapHandle);
            }
        }
        catch
        {
            return null;
        }
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void SHCreateItemFromParsingName(
        [MarshalAs(UnmanagedType.LPWStr)] string path,
        IntPtr bindingContext,
        ref Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out IShellItemImageFactory shellItemImageFactory);

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteObject(IntPtr objectHandle);

    [ComImport]
    [Guid("bcc18b79-ba16-442f-80c4-8a59c30c463b")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItemImageFactory
    {
        void GetImage(NativeSize size, ShellItemImageFactoryOptions flags, out IntPtr bitmapHandle);
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeSize
    {
        public int Width;
        public int Height;
    }

    [Flags]
    private enum ShellItemImageFactoryOptions
    {
        BiggerSizeOk = 0x1,
        ThumbnailOnly = 0x8
    }
}
