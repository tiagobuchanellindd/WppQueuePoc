namespace WppQueuePoc.Models;

/// <summary>
/// Snapshot de ticket padrão e capacidades da fila para diagnóstico funcional.
/// </summary>
public sealed record PrintTicketInfoResult(
    /// <summary>
    /// Nome da fila consultada.
    /// </summary>
    string QueueName,

    /// <summary>
    /// Indica se a leitura foi possível no runtime/ambiente atual.
    /// </summary>
    bool Available,

    /// <summary>
    /// Mensagem de contexto técnico do resultado.
    /// </summary>
    string Details,

    /// <summary>
    /// Atributos/capacidades extraídos do ticket e do driver.
    /// </summary>
    IReadOnlyDictionary<string, string> Attributes);
