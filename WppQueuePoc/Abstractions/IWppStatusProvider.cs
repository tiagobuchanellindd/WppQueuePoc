using WppQueuePoc.Models;

namespace WppQueuePoc.Abstractions;

/// <summary>
/// Fornece o estado global do Windows Protected Print (WPP) para o fluxo de negócio.
/// </summary>
public interface IWppStatusProvider
{
    /// <summary>
    /// Obtém o status global do WPP com a evidência técnica usada na decisão.
    /// </summary>
    WppStatusResult GetWppStatus();
}
