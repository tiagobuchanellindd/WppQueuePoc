namespace WppQueuePoc.Models;

/// <summary>
/// Evidências retornadas por APMON/GetAPPortInfo para classificação de protocolo.
/// </summary>
public sealed record ApPortData(
    /// <summary>
    /// Versão da estrutura de retorno.
    /// </summary>
    uint Version,

    /// <summary>
    /// Protocolo detectado (ex.: 1=WSD, 2=IPP).
    /// </summary>
    uint Protocol,

    /// <summary>
    /// URL do dispositivo/serviço quando fornecida.
    /// </summary>
    string DeviceOrServiceUrl);
