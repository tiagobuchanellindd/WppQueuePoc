namespace WppQueuePoc.Models;

/// <summary>
/// Dados resumidos de uma fila para listagem operacional.
/// </summary>
public sealed record QueueInfo(
    /// <summary>
    /// Nome da fila.
    /// </summary>
    string Name,

    /// <summary>
    /// Nome da porta usada pela fila.
    /// </summary>
    string PortName,

    /// <summary>
    /// Nome do driver associado.
    /// </summary>
    string DriverName,

    /// <summary>
    /// Indica se a fila está compartilhada.
    /// </summary>
    bool Shared,

    /// <summary>
    /// Comentário administrativo da fila.
    /// </summary>
    string Comment,

    /// <summary>
    /// Localização informada para apoio operacional.
    /// </summary>
    string Location);
