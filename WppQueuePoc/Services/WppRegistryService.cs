using System.Runtime.Versioning;
using Microsoft.Win32;
using WppQueuePoc.Abstractions;
using WppQueuePoc.Models;

namespace WppQueuePoc.Services;

/// <summary>
/// Implementa a consulta do estado global do Windows Protected Print (WPP)
/// lendo valores de política no Registro do Windows (HKLM).
///
/// A classe percorre uma lista de "probes" (caminho + nome de valor), aplica
/// validações de existência, tipo e mapeamento semântico, e converte o resultado
/// em um <see cref="WppStatusResult"/> para consumo da aplicação.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class WppRegistryService : IWppStatusProvider
{
    private static readonly PolicyStatusCheck[] StatusChecks =
    [
        //Acessar via Editor de Registro (ou regedit) => Computador\HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP
        new(
            @"SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP",
            "WindowsProtectedPrintGroupPolicyState",
            EnabledWhen: 1,
            DisabledWhen: 0),
        new(
            @"SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP",
            "WindowsProtectedPrintMode",
            EnabledWhen: 1,
            DisabledWhen: 0)
    ];

    /// <summary>
    /// Resolve o estado final do WPP a partir de chaves de política conhecidas.
    ///
    /// A rotina executa uma estratégia de fallback: para cada probe, tenta abrir a
    /// subchave, ler o valor e convertê-lo para inteiro. Em seguida, compara o
    /// número lido com os valores esperados de habilitado/desabilitado.
    ///
    /// Regras de decisão:
    /// - Retorna <see cref="WppStatus.Enabled"/> imediatamente ao encontrar um valor
    ///   mapeado como habilitado.
    /// - Guarda um candidato <see cref="WppStatus.Disabled"/> e continua procurando,
    ///   pois um probe posterior ainda pode indicar habilitado.
    /// - Retorna <see cref="WppStatus.Unknown"/> quando encontra tipo não suportado
    ///   ou valor fora do mapeamento conhecido.
    /// - Se nada for encontrado nos probes, assume desabilitado por padrão.
    /// </summary>
    public WppStatusResult GetWppStatus()
    {
        WppStatusResult? disabledCandidate = null;

        foreach (var statusCheck in StatusChecks)
        {
            using var key = Registry.LocalMachine.OpenSubKey(statusCheck.RegistryPath, writable: false);

            // Validação de existência: se a chave não existe, tenta o próximo mapeamento conhecido.
            if (key is null)
            {
                continue;
            }

            var rawValue = key.GetValue(statusCheck.PolicyValueName);
            // Validação de valor: se o valor não existe neste probe, tenta o próximo fallback.
            if (rawValue is null)
            {
                continue;
            }

            // Validação de tipo: só aceita tipos que conseguimos converter com segurança.
            if (!TryConvertToInt(rawValue, out var numericValue))
            {
                return new WppStatusResult(
                    WppStatus.Unknown,
                    $@"HKLM\{statusCheck.RegistryPath}\{statusCheck.PolicyValueName}",
                    $"Unsupported value type: {rawValue.GetType().Name}.",
                    null);
            }

            // Validação semântica: mapeia valor conhecido para "habilitado".
            if (numericValue == statusCheck.EnabledWhen)
            {
                return new WppStatusResult(
                    WppStatus.Enabled,
                    $@"HKLM\{statusCheck.RegistryPath}\{statusCheck.PolicyValueName}",
                    "Matched known enabled value.",
                    numericValue);
            }

            // Validação semântica: mapeia valor conhecido para "desabilitado".
            if (numericValue == statusCheck.DisabledWhen)
            {
                // Mantém como candidato e continua: um próximo probe pode indicar "habilitado".
                disabledCandidate ??= new WppStatusResult(
                    WppStatus.Disabled,
                    $@"HKLM\{statusCheck.RegistryPath}\{statusCheck.PolicyValueName}",
                    "Matched known disabled value.",
                    numericValue);
                continue;
            }

            // Validação final: valor fora do mapeamento conhecido, então o resultado é indeterminado.
            return new WppStatusResult(
                WppStatus.Unknown,
                $@"HKLM\{statusCheck.RegistryPath}\{statusCheck.PolicyValueName}",
                "Value found but does not match known enabled/disabled mapping.",
                numericValue);
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
    /// Tenta normalizar o valor bruto lido do Registro para inteiro.
    ///
    /// O método aceita os formatos mais comuns encontrados nesse cenário:
    /// - <see cref="int"/> (tipo esperado para DWORD no Registry).
    /// - <see cref="string"/> contendo número válido (fallback para ambientes que
    ///   persistem o valor em texto).
    ///
    /// Retorna <see langword="true"/> quando a conversão é segura, preenchendo
    /// <paramref name="result"/>; caso contrário, retorna
    /// <see langword="false"/> e define <paramref name="result"/> com o valor padrão.
    /// </summary>
    private static bool TryConvertToInt(object rawValue, out int numericValue)
    {
        // Validação de tipo nativo comum no Registry.
        if (rawValue is int integerValue)
        {
            numericValue = integerValue;
            return true;
        }

        // Validação de fallback: alguns ambientes podem armazenar como texto numérico.
        if (rawValue is string stringValue && int.TryParse(stringValue, out var parsedValue))
        {
            numericValue = parsedValue;
            return true;
        }

        numericValue = default;
        return false;
    }

    private sealed record PolicyStatusCheck(string RegistryPath, string PolicyValueName, int EnabledWhen, int DisabledWhen);
}
