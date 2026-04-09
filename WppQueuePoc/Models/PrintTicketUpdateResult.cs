namespace WppQueuePoc.Models;

/// <summary>
/// Resultado da tentativa de atualização de ticket (default/user).
/// </summary>
public sealed record PrintTicketUpdateResult(
    string QueueName,
    string Scope,
    bool Applied,
    string Details,
    IReadOnlyDictionary<string, string> Requested,
    IReadOnlyDictionary<string, string> AppliedValues);
