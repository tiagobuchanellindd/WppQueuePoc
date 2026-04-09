using System.Runtime.Versioning;
using Microsoft.Win32;
using WppQueuePoc.Abstractions;
using WppQueuePoc.Models;

namespace WppQueuePoc.Services;

/// <summary>
/// Implementação de status global WPP baseada em Registry de políticas do Windows.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class WppRegistryService : IWppStatusProvider
{
    private static readonly RegistryProbe[] Probes =
    [
        //Acessar via Editor de Registro (ou regedit) => Computador\HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP
        new(
            @"SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP",
            "WindowsProtectedPrintGroupPolicyState",
            EnabledValue: 1,
            DisabledValue: 0),
        new(
            @"SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP",
            "WindowsProtectedPrintMode",
            EnabledValue: 1,
            DisabledValue: 0)
    ];

    /// <summary>
    /// Resolve o estado global WPP a partir das chaves conhecidas de política.
    /// </summary>
    public WppStatusResult GetWppStatus()
    {
        WppStatusResult? disabledCandidate = null;

        foreach (var probe in Probes)
        {
            using var key = Registry.LocalMachine.OpenSubKey(probe.Path, writable: false);
            // Validação de existência: se a chave não existe, tenta o próximo mapeamento conhecido.
            if (key is null)
            {
                continue;
            }

            var raw = key.GetValue(probe.ValueName);
            // Validação de valor: se o valor não existe neste probe, tenta o próximo fallback.
            if (raw is null)
            {
                continue;
            }

            // Validação de tipo: só aceitamos tipos que conseguimos converter com segurança.
            if (!TryConvertToInt(raw, out var value))
            {
                return new WppStatusResult(
                    WppStatus.Unknown,
                    $@"HKLM\{probe.Path}\{probe.ValueName}",
                    $"Unsupported value type: {raw.GetType().Name}.",
                    null);
            }

            // Validação semântica: mapeia valor conhecido para "habilitado".
            if (value == probe.EnabledValue)
            {
                return new WppStatusResult(
                    WppStatus.Enabled,
                    $@"HKLM\{probe.Path}\{probe.ValueName}",
                    "Matched known enabled value.",
                    value);
            }

            // Validação semântica: mapeia valor conhecido para "desabilitado".
            if (value == probe.DisabledValue)
            {
                // Mantém como candidato e continua: um próximo probe pode indicar "habilitado".
                disabledCandidate ??= new WppStatusResult(
                    WppStatus.Disabled,
                    $@"HKLM\{probe.Path}\{probe.ValueName}",
                    "Matched known disabled value.",
                    value);
                continue;
            }

            // Validação final: valor fora do mapeamento conhecido, então o resultado é indeterminado.
            return new WppStatusResult(
                WppStatus.Unknown,
                $@"HKLM\{probe.Path}\{probe.ValueName}",
                "Value found but does not match known enabled/disabled mapping.",
                value);
        }

        if (disabledCandidate is not null)
        {
            return disabledCandidate;
        }

        return new WppStatusResult(
            WppStatus.Disabled,
            "No known WPP Registry value found.",
            "No known enabled/disabled value was found in mapped probes. Treated as disabled by default.",
            null);
    }

    /// <summary>
    /// Tenta normalizar o valor lido do Registry para inteiro.
    /// </summary>
    private static bool TryConvertToInt(object value, out int result)
    {
        // Validação de tipo nativo comum no Registry.
        if (value is int i)
        {
            result = i;
            return true;
        }

        // Validação de fallback: alguns ambientes podem armazenar como texto numérico.
        if (value is string s && int.TryParse(s, out var parsed))
        {
            result = parsed;
            return true;
        }

        result = default;
        return false;
    }

    private sealed record RegistryProbe(string Path, string ValueName, int EnabledValue, int DisabledValue);
}
