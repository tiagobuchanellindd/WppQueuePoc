using WppQueuePoc.Models;

namespace WppQueuePoc.Abstractions;

/// <summary>
/// Contrato das operações de administração de filas usadas na POC.
/// </summary>
public interface IPrintSpoolerService
{
    /// <summary>
    /// Tenta criar uma porta WSD no monitor de portas do Windows.
    /// </summary>
    void AddWsdPort(string portName);

    /// <summary>
    /// Cria uma fila de impressão com driver, porta e parâmetros de processamento.
    /// </summary>
    void CreateQueue(string queueName, string driverName, string portName, string printProcessor, string dataType, string comment, string location);

    /// <summary>
    /// Lista as filas instaladas para suporte operacional e troubleshooting.
    /// </summary>
    IReadOnlyList<QueueInfo> ListQueues();

    /// <summary>
    /// Lista as portas disponíveis no spooler (ex.: WSD, IPP e locais).
    /// </summary>
    IReadOnlyList<string> ListPorts();

    /// <summary>
    /// Lista os drivers instalados para criação/atualização de filas.
    /// </summary>
    IReadOnlyList<string> ListDrivers();

    /// <summary>
    /// Lista os processadores de impressão disponíveis no host.
    /// </summary>
    IReadOnlyList<string> ListPrintProcessors();

    /// <summary>
    /// Lista os datatypes disponíveis para um processador de impressão.
    /// </summary>
    IReadOnlyList<string> ListDataTypes(string printProcessor);

    /// <summary>
    /// Atualiza metadados e/ou propriedades técnicas de uma fila existente.
    /// </summary>
    void UpdateQueue(string queueName, string? newQueueName, string? newDriverName, string? newPortName, string? comment, string? location);

    /// <summary>
    /// Remove uma fila existente do spooler.
    /// </summary>
    void DeleteQueue(string queueName);

    /// <summary>
    /// Classifica a fila quanto à probabilidade de operação sob WPP com base em evidências.
    /// </summary>
    QueueInspectionResult InspectQueue(string queueName);
}
