using System.IO;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using Point = System.Windows.Point;

namespace PdfMergeTool.Services;

internal sealed class NativeFileDropTarget : IDisposable
{
    private const short CfHdrop = 15;
    private const short CfUnicodeText = 13;
    private const int DropEffectNone = 0;
    private const int DropEffectCopy = 1;
    private const int DragDropAlreadyRegistered = unchecked((int)0x80040101);
    private const string PageTransferTextPrefix = "PDFMERGETOOL_PAGES:";
    private static readonly string[] SupportedFileExtensions =
    [
        ".pdf",
        ".jpg",
        ".jpeg",
        ".png",
        ".bmp",
        ".gif",
        ".tif",
        ".tiff"
    ];

    private readonly IntPtr _hwnd;
    private readonly OleDropTarget _target;
    private bool _registered;

    public NativeFileDropTarget(
        IntPtr hwnd,
        Action<IReadOnlyList<string>, Point> filesDragOver,
        Action<IReadOnlyList<string>, Point> filesDropped,
        Action<string, Point> pageTransferDragOver,
        Action pageTransferDragLeave,
        Action<string, Point> pageTransferDropped)
    {
        _hwnd = hwnd;
        _target = new OleDropTarget(
            filesDragOver,
            filesDropped,
            pageTransferDragOver,
            pageTransferDragLeave,
            pageTransferDropped);
    }

    public void Register()
    {
        if (_hwnd == IntPtr.Zero || _registered)
        {
            return;
        }

        var result = RegisterDragDrop(_hwnd, _target);
        if (result == DragDropAlreadyRegistered)
        {
            RevokeDragDrop(_hwnd);
            result = RegisterDragDrop(_hwnd, _target);
        }

        Marshal.ThrowExceptionForHR(result);
        _registered = true;
    }

    public static IReadOnlyList<IntPtr> GetWindowAndDescendants(IntPtr rootHwnd)
    {
        if (rootHwnd == IntPtr.Zero)
        {
            return [];
        }

        var handles = new List<IntPtr> { rootHwnd };
        EnumChildWindows(rootHwnd, (hwnd, _) =>
        {
            handles.Add(hwnd);
            return true;
        }, IntPtr.Zero);

        return handles.Distinct().ToList();
    }

    public void Dispose()
    {
        if (!_registered)
        {
            return;
        }

        RevokeDragDrop(_hwnd);
        _registered = false;
    }

    [DllImport("ole32.dll")]
    private static extern int RegisterDragDrop(IntPtr hwnd, IOleDropTarget dropTarget);

    [DllImport("ole32.dll")]
    private static extern int RevokeDragDrop(IntPtr hwnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc enumFunc, IntPtr lParam);

    [DllImport("ole32.dll")]
    private static extern void ReleaseStgMedium(ref ComTypes.STGMEDIUM medium);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern uint DragQueryFile(IntPtr hDrop, uint fileIndex, char[]? fileName, uint fileNameSize);

    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalLock(IntPtr hMem);

    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GlobalUnlock(IntPtr hMem);

    private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    private struct PointL
    {
        public int X;
        public int Y;
    }

    [ComImport]
    [Guid("00000122-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleDropTarget
    {
        [PreserveSig]
        int DragEnter(ComTypes.IDataObject dataObject, int keyState, PointL point, ref int effect);

        [PreserveSig]
        int DragOver(int keyState, PointL point, ref int effect);

        [PreserveSig]
        int DragLeave();

        [PreserveSig]
        int Drop(ComTypes.IDataObject dataObject, int keyState, PointL point, ref int effect);
    }

    [ComVisible(true)]
    private sealed class OleDropTarget : IOleDropTarget
    {
        private readonly Action<IReadOnlyList<string>, Point> _filesDragOver;
        private readonly Action<IReadOnlyList<string>, Point> _filesDropped;
        private readonly Action<string, Point> _pageTransferDragOver;
        private readonly Action _pageTransferDragLeave;
        private readonly Action<string, Point> _pageTransferDropped;
        private string? _currentPageTransferText;
        private IReadOnlyList<string> _currentFilePaths = [];

        public OleDropTarget(
            Action<IReadOnlyList<string>, Point> filesDragOver,
            Action<IReadOnlyList<string>, Point> filesDropped,
            Action<string, Point> pageTransferDragOver,
            Action pageTransferDragLeave,
            Action<string, Point> pageTransferDropped)
        {
            _filesDragOver = filesDragOver;
            _filesDropped = filesDropped;
            _pageTransferDragOver = pageTransferDragOver;
            _pageTransferDragLeave = pageTransferDragLeave;
            _pageTransferDropped = pageTransferDropped;
        }

        public int DragEnter(ComTypes.IDataObject dataObject, int keyState, PointL point, ref int effect)
        {
            _currentPageTransferText = TryReadPageTransferText(dataObject);
            _currentFilePaths = string.IsNullOrWhiteSpace(_currentPageTransferText)
                ? TryReadSupportedPaths(dataObject)
                : [];
            UpdateEffectAndIndicator(dataObject, point, ref effect);
            return 0;
        }

        public int DragOver(int keyState, PointL point, ref int effect)
        {
            if (!string.IsNullOrWhiteSpace(_currentPageTransferText))
            {
                effect = DropEffectCopy;
                _pageTransferDragOver(_currentPageTransferText, ToPoint(point));
                return 0;
            }

            if (_currentFilePaths.Count > 0)
            {
                effect = DropEffectCopy;
                _filesDragOver(_currentFilePaths, ToPoint(point));
                return 0;
            }

            effect = DropEffectNone;
            return 0;
        }

        public int DragLeave()
        {
            _currentPageTransferText = null;
            _currentFilePaths = [];
            _pageTransferDragLeave();
            return 0;
        }

        public int Drop(ComTypes.IDataObject dataObject, int keyState, PointL point, ref int effect)
        {
            var pageTransferText = _currentPageTransferText ?? TryReadPageTransferText(dataObject);
            if (!string.IsNullOrWhiteSpace(pageTransferText))
            {
                effect = DropEffectCopy;
                _pageTransferDropped(pageTransferText, ToPoint(point));
                _currentPageTransferText = null;
                _currentFilePaths = [];
                return 0;
            }

            var filePaths = _currentFilePaths.Count > 0 ? _currentFilePaths : TryReadSupportedPaths(dataObject);
            if (filePaths.Count > 0)
            {
                effect = DropEffectCopy;
                _filesDropped(filePaths, ToPoint(point));
                _currentPageTransferText = null;
                _currentFilePaths = [];
                return 0;
            }

            effect = DropEffectNone;
            _currentPageTransferText = null;
            _currentFilePaths = [];
            return 0;
        }

        private void UpdateEffectAndIndicator(ComTypes.IDataObject dataObject, PointL point, ref int effect)
        {
            if (!string.IsNullOrWhiteSpace(_currentPageTransferText))
            {
                effect = DropEffectCopy;
                _pageTransferDragOver(_currentPageTransferText, ToPoint(point));
                return;
            }

            _currentFilePaths = TryReadSupportedPaths(dataObject);
            if (_currentFilePaths.Count > 0)
            {
                effect = DropEffectCopy;
                _filesDragOver(_currentFilePaths, ToPoint(point));
                return;
            }

            effect = DropEffectNone;
        }

        private static Point ToPoint(PointL point) => new(point.X, point.Y);

        private static List<string> TryReadSupportedPaths(ComTypes.IDataObject dataObject)
        {
            var paths = new List<string>();
            var format = CreateFormat(CfHdrop);
            ComTypes.STGMEDIUM medium = default;

            try
            {
                dataObject.GetData(ref format, out medium);
                var count = DragQueryFile(medium.unionmember, uint.MaxValue, null, 0);
                for (uint index = 0; index < count; index++)
                {
                    var length = DragQueryFile(medium.unionmember, index, null, 0);
                    if (length == 0)
                    {
                        continue;
                    }

                    var buffer = new char[length + 1];
                    DragQueryFile(medium.unionmember, index, buffer, (uint)buffer.Length);
                    var path = new string(buffer).TrimEnd('\0');
                    if (File.Exists(path) && IsSupportedFile(path))
                    {
                        paths.Add(Path.GetFullPath(path));
                    }
                }
            }
            catch (COMException)
            {
                return [];
            }
            finally
            {
                if (medium.unionmember != IntPtr.Zero)
                {
                    ReleaseStgMedium(ref medium);
                }
            }

            return paths;
        }

        private static bool IsSupportedFile(string path)
        {
            var extension = Path.GetExtension(path);
            return SupportedFileExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase);
        }

        private static string? TryReadPageTransferText(ComTypes.IDataObject dataObject)
        {
            var format = CreateFormat(CfUnicodeText);
            ComTypes.STGMEDIUM medium = default;

            try
            {
                dataObject.GetData(ref format, out medium);
                var pointer = GlobalLock(medium.unionmember);
                if (pointer == IntPtr.Zero)
                {
                    return null;
                }

                try
                {
                    var text = Marshal.PtrToStringUni(pointer);
                    return text?.StartsWith(PageTransferTextPrefix, StringComparison.Ordinal) == true
                        ? text[PageTransferTextPrefix.Length..]
                        : null;
                }
                finally
                {
                    GlobalUnlock(medium.unionmember);
                }
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                if (medium.unionmember != IntPtr.Zero)
                {
                    ReleaseStgMedium(ref medium);
                }
            }
        }

        private static ComTypes.FORMATETC CreateFormat(short format) => new()
        {
            cfFormat = format,
            dwAspect = ComTypes.DVASPECT.DVASPECT_CONTENT,
            lindex = -1,
            ptd = IntPtr.Zero,
            tymed = ComTypes.TYMED.TYMED_HGLOBAL
        };
    }
}
