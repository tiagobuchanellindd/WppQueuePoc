using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace WppQueuePoc.Interop;

/// <summary>
/// Assinaturas P/Invoke e estruturas nativas do <c>winspool.drv</c>.
/// Isola a fronteira de interoperabilidade da aplicação.
/// </summary>
[SupportedOSPlatform("windows")]
internal static class NativeMethods
{
    public const uint PRINTER_ENUM_LOCAL = 0x00000002;
    public const uint PRINTER_ENUM_CONNECTIONS = 0x00000004;
    public const uint PRINTER_ATTRIBUTE_QUEUED = 0x00000001;
    public const uint PRINTER_ATTRIBUTE_SHARED = 0x00000008;
    public const uint SERVER_ACCESS_ADMINISTER = 0x00000001;
    public const uint PRINTER_ACCESS_USE = 0x00000008;
    public const uint PRINTER_ACCESS_ADMINISTER = 0x00000004;
    public const uint PRINTER_ALL_ACCESS = 0x000F000C;

    public const int ERROR_SUCCESS = 0;
    public const int ERROR_ACCESS_DENIED = 5;
    public const int ERROR_NOT_SUPPORTED = 50;
    public const int ERROR_INSUFFICIENT_BUFFER = 122;

    /// <summary>
    /// Parâmetros padrão para operações com <c>OpenPrinter</c>.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct PRINTER_DEFAULTS
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pDatatype;
        public IntPtr pDevMode;
        public uint DesiredAccess;
    }

    /// <summary>
    /// Estrutura de metadados completos de fila (nível 2).
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct PRINTER_INFO_2
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pServerName;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pPrinterName;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pShareName;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pPortName;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pDriverName;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pComment;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pLocation;
        public IntPtr pDevMode;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pSepFile;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pPrintProcessor;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pDatatype;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pParameters;
        public IntPtr pSecurityDescriptor;
        public uint Attributes;
        public uint Priority;
        public uint DefaultPriority;
        public uint StartTime;
        public uint UntilTime;
        public uint Status;
        public uint cJobs;
        public uint AveragePPM;
    }

    /// <summary>
    /// Estrutura de porta simples (nível 1).
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct PORT_INFO_1
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pName;
    }

    /// <summary>
    /// Estrutura de driver (nível 2).
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DRIVER_INFO_2
    {
        public uint cVersion;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pName;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pEnvironment;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pDriverPath;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pDataFile;
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pConfigFile;
    }

    /// <summary>
    /// Estrutura de processador de impressão (nível 1).
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct PRINTPROCESSOR_INFO_1
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pName;
    }

    /// <summary>
    /// Estrutura de datatype de processador (nível 1).
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DATATYPES_INFO_1
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pName;
    }

    /// <summary>
    /// Estrutura retornada por GetAPPortInfo para evidências de protocolo moderno.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct AP_PORT_DATA_1
    {
        public uint Version;
        public uint Protocol;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string DeviceOrServiceUrl;
    }

    /// <summary>
    /// Abre um handle de impressora, monitor ou Xcv endpoint.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, ref PRINTER_DEFAULTS pDefault);

    /// <summary>
    /// Fecha handle previamente aberto por <c>OpenPrinter</c>.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true)]
    public static extern bool ClosePrinter(IntPtr hPrinter);

    /// <summary>
    /// Executa comando Xcv com payload em bytes.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool XcvData(
        IntPtr hXcv,
        string pszDataName,
        byte[] pInputData,
        int cbInputData,
        IntPtr pOutputData,
        int cbOutputData,
        out uint pcbOutputNeeded,
        out uint pdwStatus);

    /// <summary>
    /// Executa comando Xcv com payload em ponteiro nativo.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool XcvData(
        IntPtr hXcv,
        string pszDataName,
        IntPtr pInputData,
        uint cbInputData,
        IntPtr pOutputData,
        uint cbOutputData,
        out uint pcbOutputNeeded,
        out uint pdwStatus);

    /// <summary>
    /// Cria fila de impressão.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr AddPrinter(string? pName, uint Level, IntPtr pPrinter);

    /// <summary>
    /// Atualiza propriedades da fila.
    /// </summary>
    [DllImport("winspool.drv", EntryPoint = "SetPrinterW", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool SetPrinter(IntPtr hPrinter, uint Level, IntPtr pPrinter, uint Command);

    /// <summary>
    /// Exclui fila de impressão.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true)]
    public static extern bool DeletePrinter(IntPtr hPrinter);

    /// <summary>
    /// Enumera filas no spooler.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool EnumPrinters(
        uint Flags,
        string? Name,
        uint Level,
        IntPtr pPrinterEnum,
        uint cbBuf,
        out uint pcbNeeded,
        out uint pcReturned);

    /// <summary>
    /// Obtém dados da fila por nível de informação.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool GetPrinter(
        IntPtr hPrinter,
        uint Level,
        IntPtr pPrinter,
        uint cbBuf,
        out uint pcbNeeded);

    /// <summary>
    /// Enumera portas de impressão.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool EnumPorts(
        string? pName,
        uint Level,
        IntPtr pPortInfo,
        uint cbBuf,
        out uint pcbNeeded,
        out uint pcReturned);

    /// <summary>
    /// Enumera drivers instalados.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool EnumPrinterDrivers(
        string? pName,
        string? pEnvironment,
        uint Level,
        IntPtr pDriverInfo,
        uint cbBuf,
        out uint pcbNeeded,
        out uint pcReturned);

    /// <summary>
    /// Enumera processadores de impressão.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool EnumPrintProcessors(
        string? pName,
        string? pEnvironment,
        uint Level,
        IntPtr pPrintProcessorInfo,
        uint cbBuf,
        out uint pcbNeeded,
        out uint pcReturned);

    /// <summary>
    /// Enumera datatypes de um processador.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool EnumPrintProcessorDatatypes(
        string? pName,
        string pPrintProcessorName,
        uint Level,
        IntPtr pDatatypes,
        uint cbBuf,
        out uint pcbNeeded,
        out uint pcReturned);
}
