namespace WppQueuePoc.Models;

/// <summary>
/// Representa o estado global do Windows Protected Print (WPP).
/// </summary>
public enum WppStatus
{
    /// <summary>
    /// Política global de WPP habilitada.
    /// </summary>
    Enabled,

    /// <summary>
    /// Política global de WPP desabilitada.
    /// </summary>
    Disabled,

    /// <summary>
    /// Não foi possível determinar o estado com segurança.
    /// </summary>
    Unknown
}
