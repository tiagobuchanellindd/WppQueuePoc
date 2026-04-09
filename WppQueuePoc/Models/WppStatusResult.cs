namespace WppQueuePoc.Models;

/// <summary>
/// Resultado da consulta de status global WPP com evidências técnicas.
/// </summary>
public sealed record WppStatusResult(
    /// <summary>
    /// Status final inferido.
    /// </summary>
    WppStatus Status,

    /// <summary>
    /// Fonte da informação (ex.: chave/valor de Registry).
    /// </summary>
    string Source,

    /// <summary>
    /// Explicação técnica usada na decisão.
    /// </summary>
    string Details,

    /// <summary>
    /// Valor bruto encontrado na fonte, quando disponível.
    /// </summary>
    int? RawValue);
