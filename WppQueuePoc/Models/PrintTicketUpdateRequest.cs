namespace WppQueuePoc.Models;

/// <summary>
/// Solicitação de atualização parcial do Print Ticket.
/// Valores devem ser nomes de enums do System.Printing (ex.: TwoSidedLongEdge, Color, Landscape).
/// </summary>
public sealed record PrintTicketUpdateRequest(
    string? Duplexing,
    string? OutputColor,
    string? PageOrientation);
