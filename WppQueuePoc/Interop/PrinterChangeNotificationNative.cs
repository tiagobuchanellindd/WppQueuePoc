using System;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace WppQueuePoc.Interop;

[SupportedOSPlatform("windows")]
internal static class PrinterChangeNotificationNative
{
    // Win32 API: FindFirstPrinterChangeNotification
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr FindFirstPrinterChangeNotification(
        IntPtr hPrinter,
        int fdwFilter,
        int fdwOptions,
        IntPtr pPrinterNotifyOptions);

    // Win32 API: FindNextPrinterChangeNotification
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool FindNextPrinterChangeNotification(
        IntPtr hChange,
        out int pdwChange,
        IntPtr pPrinterNotifyOptions,
        out IntPtr ppPrinterNotifyInfo);

    // Win32 API: FindClosePrinterChangeNotification
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool FindClosePrinterChangeNotification(IntPtr hChange);

    // Win32 API: WaitForSingleObject
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern int WaitForSingleObject(IntPtr handle, int milliseconds);

    // PRINTER_CHANGE_* flags
    public const int PRINTER_CHANGE_PRINTER = 0x000000FF;
    public const int PRINTER_CHANGE_ADD_PRINTER = 0x00000001;
    public const int PRINTER_CHANGE_SET_PRINTER = 0x00000002;
    public const int PRINTER_CHANGE_DELETE_PRINTER = 0x00000004;
    public const int PRINTER_CHANGE_FAILED_CONNECTION_PRINTER = 0x00000008;

    public const int WAIT_OBJECT_0 = 0;
    public const int WAIT_TIMEOUT = 0x102;
    public const int WAIT_FAILED = unchecked((int)0xFFFFFFFF);
}
