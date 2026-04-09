using WppQueuePoc.Models;

namespace WppQueuePoc.Abstractions;

/// <summary>
/// Contrato para leitura de informações de Print Ticket, útil para diagnóstico funcional da fila.
/// </summary>
public interface IPrintTicketService
{
    /// <summary>
    /// Lê o ticket padrão e capacidades básicas da fila informada.
    /// </summary>
    PrintTicketInfoResult GetDefaultTicketInfo(string queueName);

    /// <summary>
    /// Atualiza propriedades no ticket padrão da fila.
    /// </summary>
    PrintTicketUpdateResult UpdateDefaultTicket(string queueName, PrintTicketUpdateRequest request);

    /// <summary>
    /// Atualiza propriedades no ticket padrão do usuário para a fila.
    /// </summary>
    PrintTicketUpdateResult UpdateUserTicket(string queueName, PrintTicketUpdateRequest request);
    /// <summary>
    /// Lê o ticket do usuário da fila informada.
    /// </summary>
    PrintTicketInfoResult GetUserTicketInfo(string queueName);
}

