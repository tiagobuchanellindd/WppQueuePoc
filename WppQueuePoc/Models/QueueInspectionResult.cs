namespace WppQueuePoc.Models;

/// <summary>
/// Resultado de inspeção de fila com sinais de negócio e técnicos para WPP.
/// </summary>
public sealed record QueueInspectionResult(
    /// <summary>
    /// Nome da fila inspecionada.
    /// </summary>
    string QueueName,

    /// <summary>
    /// Porta associada à fila.
    /// </summary>
    string PortName,

    /// <summary>
    /// Status global WPP usado na classificação.
    /// </summary>
    WppStatus GlobalWppStatus,

    /// <summary>
    /// Resultado heurístico final.
    /// </summary>
    WppQueueClassification Classification,

    /// <summary>
    /// Detalhes e evidências coletadas no processo.
    /// </summary>
    string Details);
