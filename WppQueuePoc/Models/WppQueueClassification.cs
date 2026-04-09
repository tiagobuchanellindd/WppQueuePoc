namespace WppQueuePoc.Models;

/// <summary>
/// Classificação heurística da fila em relação ao cenário WPP.
/// </summary>
public enum WppQueueClassification
{
    /// <summary>
    /// Evidências indicam alta probabilidade de operação sob WPP.
    /// </summary>
    LikelyWpp,

    /// <summary>
    /// Evidências indicam alta probabilidade de não operação sob WPP.
    /// </summary>
    LikelyNotWpp,

    /// <summary>
    /// Evidências insuficientes ou conflitantes para concluir.
    /// </summary>
    Indeterminate
}
