using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace WppQueuePoc.Interop;

/// <summary>
/// Centraliza assinaturas P/Invoke, estruturas Win32 e constantes do subsistema
/// de impressao (<c>winspool.drv</c>).
///
/// Esta classe define a fronteira de interoperabilidade da aplicacao:
/// - Expõe chamadas nativas usadas pelos services.
/// - Declara layouts de memoria para marshal entre mundo gerenciado e nativo.
/// - Mantem codigos/flags em um unico ponto para reduzir inconsistencias.
///
/// Como guia de uso, quase todas as APIs de enumeracao seguem o padrao
/// "duas chamadas": primeira para descobrir tamanho de buffer e segunda para
/// preencher os dados efetivos.
/// </summary>
[SupportedOSPlatform("windows")]
internal static class NativeMethods
{
    /// <summary>
    /// Enumera filas locais instaladas no host atual.
    /// Usado em <c>EnumPrinters</c> para descobrir impressoras locais.
    /// </summary>
    public const uint PRINTER_ENUM_LOCAL = 0x00000002;

    /// <summary>
    /// Enumera conexoes de impressora (ex.: filas conectadas de outros servidores).
    /// Usado em conjunto com <c>PRINTER_ENUM_LOCAL</c> para visao consolidada.
    /// </summary>
    public const uint PRINTER_ENUM_CONNECTIONS = 0x00000004;

    /// <summary>
    /// Define atributo de fila enfileirada (spooling habilitado).
    /// Aplicado na criacao de fila via <c>PRINTER_INFO_2.Attributes</c>.
    /// </summary>
    public const uint PRINTER_ATTRIBUTE_QUEUED = 0x00000001;

    /// <summary>
    /// Indica que a fila esta compartilhada para uso em rede.
    /// Lido de <c>PRINTER_INFO_2.Attributes</c> durante inspecao/listagem.
    /// </summary>
    public const uint PRINTER_ATTRIBUTE_SHARED = 0x00000008;

    /// <summary>
    /// Nivel de acesso administrativo a servidor/monitor de impressao.
    /// Necessario para operacoes Xcv, como criacao de portas.
    /// </summary>
    public const uint SERVER_ACCESS_ADMINISTER = 0x00000001;

    /// <summary>
    /// Nivel minimo para uso/consulta basica de uma fila.
    /// Tipicamente suficiente para leitura de propriedades.
    /// </summary>
    public const uint PRINTER_ACCESS_USE = 0x00000008;

    /// <summary>
    /// Nivel de acesso para administrar configuracoes da fila.
    /// Necessario para operacoes como <c>SetPrinter</c>.
    /// </summary>
    public const uint PRINTER_ACCESS_ADMINISTER = 0x00000004;

    /// <summary>
    /// Mascara de acesso completo para a fila.
    /// Usada em cenarios destrutivos, como <c>DeletePrinter</c>.
    /// </summary>
    public const uint PRINTER_ALL_ACCESS = 0x000F000C;

    /// <summary>
    /// Operacao concluida com sucesso (codigo Win32 0).
    /// </summary>
    public const int ERROR_SUCCESS = 0;

    /// <summary>
    /// Acesso negado por permissao insuficiente (codigo Win32 5).
    /// </summary>
    public const int ERROR_ACCESS_DENIED = 5;

    /// <summary>
    /// Operacao/comando nao suportado no ambiente atual (codigo Win32 50).
    /// </summary>
    public const int ERROR_NOT_SUPPORTED = 50;

    /// <summary>
    /// Buffer fornecido e menor que o necessario (codigo Win32 122).
    /// Esperado na primeira chamada de APIs de enumeracao.
    /// </summary>
    public const int ERROR_INSUFFICIENT_BUFFER = 122;

    /// <summary>
    /// Parametros padrao para chamadas de <c>OpenPrinter</c>.
    ///
    /// Define datatype opcional, ponteiro para DEVMODE (normalmente nulo neste
    /// projeto) e mascara de acesso desejada para o handle retornado.
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
    /// Estrutura <c>PRINTER_INFO_2</c> com metadados completos da fila.
    ///
    /// E a estrutura principal para criar, ler e atualizar configuracoes de
    /// impressora no spooler: nome, driver, porta, atributos, prioridade,
    /// dados de seguranca e propriedades operacionais.
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
    /// Estrutura <c>PORT_INFO_1</c> (nivel 1) para enumeracao de portas.
    ///
    /// Contem somente o nome da porta, suficiente para fluxos de selecao e
    /// validacao de existencia de porta no host.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct PORT_INFO_1
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pName;
    }

    /// <summary>
    /// Estrutura <c>DRIVER_INFO_2</c> (nivel 2) para drivers de impressao.
    ///
    /// Inclui nome, ambiente de execucao e caminhos principais de binarios
    /// associados ao driver instalado.
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
    /// Estrutura <c>PRINTPROCESSOR_INFO_1</c> (nivel 1).
    ///
    /// Exposta para enumerar os processadores de impressao registrados.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct PRINTPROCESSOR_INFO_1
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pName;
    }

    /// <summary>
    /// Estrutura <c>DATATYPES_INFO_1</c> (nivel 1).
    ///
    /// Representa cada datatype suportado por um processador de impressao.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DATATYPES_INFO_1
    {
        [MarshalAs(UnmanagedType.LPWStr)]
        public string? pName;
    }

    /// <summary>
    /// Estrutura retornada por <c>GetAPPortInfo</c> para portas APMON.
    ///
    /// Traz versao do payload, protocolo detectado (ex.: WSD/IPP) e URL fixa
    /// do dispositivo/servico, usada como evidencia complementar em heuristicas.
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
    /// Abre um handle de impressora, monitor ou endpoint Xcv.
    ///
    /// O alvo e determinado por <paramref name="pPrinterName"/> (ex.: nome da fila,
    /// ",XcvMonitor ...", ",XcvPort ...") e as permissoes por
    /// <paramref name="pDefault"/>.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, ref PRINTER_DEFAULTS pDefault);

    /// <summary>
    /// Fecha handle aberto previamente por <c>OpenPrinter</c>.
    ///
    /// Deve ser chamado em <c>finally</c> para evitar vazamento de recursos nativos.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true)]
    public static extern bool ClosePrinter(IntPtr hPrinter);

    /// <summary>
    /// Executa comando Xcv com entrada em buffer de bytes.
    ///
    /// Sobrecarga usada quando o payload e montado como array (ex.: string Unicode
    /// terminada em nulo para comandos como <c>AddPort</c>).
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
    /// Executa comando Xcv com entrada em ponteiro nativo.
    ///
    /// Sobrecarga indicada para payloads estruturados em memoria nao gerenciada
    /// e para comandos que devolvem estruturas em buffer de saida.
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
    /// Cria uma nova fila de impressao no servidor alvo.
    ///
    /// Recebe uma estrutura de impressora no nivel informado (neste projeto,
    /// nivel 2 com <c>PRINTER_INFO_2</c>) e retorna handle da fila criada.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern IntPtr AddPrinter(string? pName, uint Level, IntPtr pPrinter);

    /// <summary>
    /// Atualiza propriedades da fila com base na estrutura informada.
    ///
    /// Usada para persistir alteracoes de configuracao, tipicamente com
    /// <c>Level=2</c> e dados de <c>PRINTER_INFO_2</c>.
    /// </summary>
    [DllImport("winspool.drv", EntryPoint = "SetPrinterW", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool SetPrinter(IntPtr hPrinter, uint Level, IntPtr pPrinter, uint Command);

    /// <summary>
    /// Exclui a fila associada ao handle informado.
    ///
    /// Requer handle aberto com privilegios adequados (geralmente all access).
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true)]
    public static extern bool DeletePrinter(IntPtr hPrinter);

    /// <summary>
    /// Enumera filas no spooler conforme flags e nivel de informacao.
    ///
    /// Para <c>Level=2</c>, o buffer contem um array de <c>PRINTER_INFO_2</c>.
    /// Deve ser chamada no padrao de dupla etapa (query de tamanho + leitura).
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
    /// Obtem dados de uma fila aberta por nivel de informacao.
    ///
    /// No projeto, e usada com <c>Level=2</c> para recuperar a estrutura completa
    /// da fila. Tambem segue o padrao de duas chamadas para buffer.
    /// </summary>
    [DllImport("winspool.drv", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool GetPrinter(
        IntPtr hPrinter,
        uint Level,
        IntPtr pPrinter,
        uint cbBuf,
        out uint pcbNeeded);

    /// <summary>
    /// Enumera portas de impressao registradas.
    ///
    /// Com <c>Level=1</c>, retorna um array de <c>PORT_INFO_1</c> no buffer.
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
    /// Enumera drivers de impressao instalados.
    ///
    /// Com <c>Level=2</c>, retorna estruturas <c>DRIVER_INFO_2</c>.
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
    /// Enumera processadores de impressao disponiveis.
    ///
    /// Com <c>Level=1</c>, retorna estruturas <c>PRINTPROCESSOR_INFO_1</c>.
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
    /// Enumera datatypes suportados por um processador de impressao.
    ///
    /// Com <c>Level=1</c>, retorna estruturas <c>DATATYPES_INFO_1</c>.
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
