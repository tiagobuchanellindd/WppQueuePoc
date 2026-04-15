using WppQueuePoc.Abstractions;
using WppQueuePoc.Models;
using System;
using System.Collections.Generic;
namespace WppQueuePoc.Services;

/// <summary>
/// Service responsavel por ler e atualizar PrintTicket de filas de impressao.
///
/// A implementacao usa reflection sobre <c>System.Printing</c> para evitar
/// dependencia estatica forte da API WPF/Printing em todos os ambientes.
/// Com isso, o codigo tenta operar quando a montagem esta disponivel e,
/// quando nao estiver, retorna resultados descritivos sem quebrar o fluxo.
/// </summary>
public sealed partial class PrintTicketService : IPrintTicketService
{
    /// <summary>
    /// Le os atributos principais do <c>DefaultPrintTicket</c> de uma fila.
    ///
    /// O metodo valida disponibilidade de <c>System.Printing</c>, abre o
    /// <c>LocalPrintServer</c>, localiza a fila e extrai propriedades comuns
    /// do ticket (cor, orientacao, duplex, etc.). O retorno sempre vem em
    /// <see cref="PrintTicketInfoResult"/>, inclusive em cenarios de erro.
    /// </summary>
    public PrintTicketInfoResult GetDefaultTicketInfo(string queueName)
    {
        var ticketSettings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var printServerType = Type.GetType("System.Printing.LocalPrintServer, System.Printing", throwOnError: false);
        if (printServerType is null)
        {
            return new PrintTicketInfoResult(
                queueName,
                Available: false,
                Details: "System.Printing assembly is not available in this runtime.",
                Attributes: ticketSettings);
        }
        object? printServer = null;
        object? queue = null;
        try
        {
            printServer = CreateLocalPrintServer(printServerType, "AdministrateServer");
            if (printServer is null)
            {
                return new PrintTicketInfoResult(queueName, false, "Unable to create LocalPrintServer.", ticketSettings);
            }
            queue = GetPrintQueue(printServerType, printServer, queueName, "DefaultPrintTicket", "AdministratePrinter");
            if (queue is null)
            {
                return new PrintTicketInfoResult(queueName, false, $"Queue '{queueName}' not found.", ticketSettings);
            }
            var queueType = queue.GetType();
            var defaultTicketPropertyInfo = queueType.GetProperty("DefaultPrintTicket");
            var defaultTicket = defaultTicketPropertyInfo?.GetValue(queue);
            if (defaultTicket is null)
            {
                return new PrintTicketInfoResult(queueName, false, "DefaultPrintTicket is not available.", ticketSettings);
            }
            ReadTicketAttribute(defaultTicket, "OutputColor", ticketSettings);
            ReadTicketAttribute(defaultTicket, "PageMediaSize", ticketSettings);
            ReadTicketAttribute(defaultTicket, "PageOrientation", ticketSettings);
            ReadTicketAttribute(defaultTicket, "InputBin", ticketSettings);
            ReadTicketAttribute(defaultTicket, "Duplexing", ticketSettings);
            ReadTicketAttribute(defaultTicket, "CopyCount", ticketSettings);
            ReadTicketAttribute(defaultTicket, "Collation", ticketSettings);
            ReadTicketAttribute(defaultTicket, "Stapling", ticketSettings);
            return new PrintTicketInfoResult(
                queueName,
                Available: true,
                Details: "Default print ticket information captured.",
                Attributes: ticketSettings);
        }
        catch (Exception ex)
        {
            return new PrintTicketInfoResult(
                queueName,
                Available: false,
                Details: $"Failed to read default print ticket information: {ex.Message}",
                Attributes: ticketSettings);
        }
        finally
        {
            DisposeIfPossible(queue);
            DisposeIfPossible(printServer);
        }
    }

    /// <summary>
    /// Le os atributos principais do <c>UserPrintTicket</c> de uma fila.
    ///
    /// A logica segue o mesmo fluxo de leitura do ticket padrao, mas usa o
    /// escopo de preferencias do usuario. Isso permite comparar configuracoes
    /// efetivas por escopo (Default vs User) para diagnostico e aprendizado.
    /// </summary>
    public PrintTicketInfoResult GetUserTicketInfo(string queueName)
    {
        var ticketSettings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var printServerType = Type.GetType("System.Printing.LocalPrintServer, System.Printing", throwOnError: false);
        if (printServerType is null)
        {
            return new PrintTicketInfoResult(
                queueName,
                Available: false,
                Details: "System.Printing assembly is not available in this runtime.",
                Attributes: ticketSettings);
        }
        object? printServer = null;
        object? queue = null;
        try
        {
            printServer = CreateLocalPrintServer(printServerType, "AdministrateServer");
            if (printServer is null)
            {
                return new PrintTicketInfoResult(queueName, false, "Unable to create LocalPrintServer.", ticketSettings);
            }
            queue = GetPrintQueue(printServerType, printServer, queueName, "UserPrintTicket", "AdministratePrinter");
            if (queue is null)
            {
                return new PrintTicketInfoResult(queueName, false, $"Queue '{queueName}' not found.", ticketSettings);
            }
            var queueType = queue.GetType();
            var userTicketPropertyInfo = queueType.GetProperty("UserPrintTicket");
            var userTicket = userTicketPropertyInfo?.GetValue(queue);
            if (userTicket is null)
            {
                return new PrintTicketInfoResult(queueName, false, "UserPrintTicket is not available.", ticketSettings);
            }
            ReadTicketAttribute(userTicket, "OutputColor", ticketSettings);
            ReadTicketAttribute(userTicket, "PageMediaSize", ticketSettings);
            ReadTicketAttribute(userTicket, "PageOrientation", ticketSettings);
            ReadTicketAttribute(userTicket, "InputBin", ticketSettings);
            ReadTicketAttribute(userTicket, "Duplexing", ticketSettings);
            ReadTicketAttribute(userTicket, "CopyCount", ticketSettings);
            ReadTicketAttribute(userTicket, "Collation", ticketSettings);
            ReadTicketAttribute(userTicket, "Stapling", ticketSettings);
            return new PrintTicketInfoResult(
                queueName,
                Available: true,
                Details: "User print ticket information captured.",
                Attributes: ticketSettings);
        }
        catch (Exception ex)
        {
            return new PrintTicketInfoResult(
                queueName,
                Available: false,
                Details: $"Failed to read user print ticket information: {ex.Message}",
                Attributes: ticketSettings);
        }
        finally
        {
            DisposeIfPossible(queue);
            DisposeIfPossible(printServer);
        }
    }

    /// <summary>
    /// Atualiza o <c>DefaultPrintTicket</c> da fila com os valores solicitados.
    ///
    /// Encaminha para o fluxo interno comum de atualizacao, fixando o escopo
    /// "Default" e a propriedade alvo do ticket.
    /// </summary>
    public PrintTicketUpdateResult UpdateDefaultTicket(string queueName, PrintTicketUpdateRequest request)
    => UpdatePrintTicketInternal(queueName, request, "DefaultPrintTicket", "Default");

    /// <summary>
    /// Atualiza o <c>UserPrintTicket</c> da fila com os valores solicitados.
    ///
    /// Encaminha para o fluxo interno comum de atualizacao, fixando o escopo
    /// "User" e a propriedade alvo do ticket.
    /// </summary>
    public PrintTicketUpdateResult UpdateUserTicket(string queueName, PrintTicketUpdateRequest request)
        => UpdatePrintTicketInternal(queueName, request, "UserPrintTicket", "User");

    /// <summary>
    /// Executa o fluxo central de atualizacao de PrintTicket para um escopo.
    ///
    /// O metodo: (1) prepara os atributos solicitados, (2) abre servidor/fila,
    /// (3) le o ticket alvo, (4) aplica apenas diferencas detectadas,
    /// (5) reatribui o ticket e faz <c>Commit</c>, e (6) retorna o que foi
    /// solicitado versus o que ficou aplicado. Erros sao capturados e
    /// transformados em mensagem explicativa no resultado.
    /// </summary>
    private PrintTicketUpdateResult UpdatePrintTicketInternal(
    string queueName,
    PrintTicketUpdateRequest request,
    string ticketPropertyName,
    string ticketScope)
    {
        var requestedSettings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        void TryAdd(string settingName, string? settingValue)
        {
            if (!string.IsNullOrWhiteSpace(settingValue))
                requestedSettings[settingName] = settingValue.Trim();
        }
        TryAdd("Duplexing", request.Duplexing);
        TryAdd("OutputColor", request.OutputColor);
        var printServerType = Type.GetType("System.Printing.LocalPrintServer, System.Printing", throwOnError: false);
        if (printServerType is null)
        {
            return new PrintTicketUpdateResult(
                queueName,
                ticketScope,
                false,
                "System.Printing not available.",
                requestedSettings,
                new Dictionary<string, string>());
        }
        object? printServer = null;
        object? queue = null;
        try
        {
            printServer = CreateLocalPrintServer(printServerType, "AdministrateServer");
            queue = GetPrintQueue(printServerType, printServer, queueName, ticketPropertyName, "AdministratePrinter");
            if (queue is null)
                return new PrintTicketUpdateResult(
                    queueName,
                    ticketScope,
                    false,
                    $"Queue '{queueName}' not found.",
                    requestedSettings,
                    new Dictionary<string, string>());
            var queueType = queue.GetType();
            var ticketPropertyInfo = queueType.GetProperty(ticketPropertyName);
            if (ticketPropertyInfo == null)
                return new PrintTicketUpdateResult(
                    queueName,
                    ticketScope,
                    false,
                    $"Print ticket property '{ticketPropertyName}' not found.",
                    requestedSettings,
                    new Dictionary<string, string>());
            var targetTicket = ticketPropertyInfo.GetValue(queue);
            if (targetTicket is null)
                return new PrintTicketUpdateResult(
                    queueName,
                    ticketScope,
                    false,
                    $"Print ticket '{ticketPropertyName}' not found (might be unsupported by this driver or printer type).",
                    requestedSettings,
                    new Dictionary<string, string>());
            // Aplica os atributos
            bool hasChanges = false;
            hasChanges |= WriteTicketAttribute(targetTicket, "Duplexing", request.Duplexing);
            hasChanges |= WriteTicketAttribute(targetTicket, "OutputColor", request.OutputColor);
            hasChanges |= WriteTicketAttribute(targetTicket, "PageOrientation", request.PageOrientation);
            string? saveErrorDetails = null;
            if (hasChanges)
            {
                try
                {
                    ticketPropertyInfo.SetValue(queue, targetTicket);
                }
                catch (Exception setTicketException)
                {
                    var error = GetInnermostMessage(setTicketException);
                    saveErrorDetails = $"Failed to set updated ticket back on '{ticketPropertyName}': {error}";
                }

                try
                {
                    var commitMethod = queueType.GetMethod("Commit");
                    commitMethod?.Invoke(queue, null);
                }
                catch (Exception commitException)
                {
                    var error = GetInnermostMessage(commitException);
                    saveErrorDetails = $"Commit failed: {error}";
                }
            }
            // Lê os valores efetivos após possível commit/aplicação
            var appliedSettings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            ReadTicketAttribute(targetTicket, "Duplexing", appliedSettings);
            ReadTicketAttribute(targetTicket, "OutputColor", appliedSettings);
            ReadTicketAttribute(targetTicket, "PageOrientation", appliedSettings);
            // Result message
            string statusMessage = hasChanges
                ? (saveErrorDetails == null ? "Ticket updated successfully." : $"Ticket update attempted. {saveErrorDetails}")
                : "No changes needed: the provided values matched current settings.";

            return new PrintTicketUpdateResult(
                queueName,
                ticketScope,
                hasChanges,
                statusMessage,
                requestedSettings,
                appliedSettings);
        }
        catch (Exception ex)
        {
            var error = GetInnermostMessage(ex);
            return new PrintTicketUpdateResult(
                queueName,
                ticketScope,
                false,
                $"Exception during PrintTicket update: {error}. This may occur if the process is not elevated and/or the account does not have 'Manage Printers' permission on this queue.",
                requestedSettings,
                new Dictionary<string, string>());
        }
        finally
        {
            DisposeIfPossible(queue);
            DisposeIfPossible(printServer);
        }
    }

    /// <summary>
    /// Le uma propriedade do ticket por reflection e grava no dicionario de saida.
    ///
    /// A leitura e best-effort: se a propriedade nao existir ou ocorrer excecao,
    /// o metodo ignora a falha para nao interromper a coleta dos demais atributos.
    /// </summary>
    private static void ReadTicketAttribute(object? ticket, string propertyName, IDictionary<string, string> settingsOutput)
    {
        try
        {
            if (ticket == null) return;
            var ticketType = ticket.GetType();
            var propertyInfo = ticketType.GetProperty(propertyName);
            if (propertyInfo != null)
            {
                var propertyValue = propertyInfo.GetValue(ticket);
                settingsOutput[propertyName] = propertyValue?.ToString() ?? "";
            }
        }
        catch { /* best effort, ignore */ }
    }

    /// <summary>
    /// Tenta escrever um atributo no ticket apenas quando houver mudanca real.
    ///
    /// Converte o valor string para o tipo da propriedade (incluindo enum e
    /// nullable), compara com o valor atual e retorna <see langword="true"/>
    /// apenas quando uma alteracao efetiva foi aplicada.
    /// </summary>
    private static bool WriteTicketAttribute(object? ticket, string propertyName, string? requestedValue)
{
    try
    {
        if (ticket == null || string.IsNullOrWhiteSpace(requestedValue))
            return false;
        var ticketType = ticket.GetType();
        var propertyInfo = ticketType.GetProperty(propertyName);
        if (propertyInfo == null || !propertyInfo.CanWrite) return false;

        var currentValue = propertyInfo.GetValue(ticket);
        var propertyType = propertyInfo.PropertyType;
        object? convertedValue = ConvertIfPossible(propertyType, requestedValue);

        bool isDifferent;
        if (currentValue is null && convertedValue is not null) isDifferent = true;
        else if (currentValue is not null && convertedValue is null) isDifferent = true;
        else if (currentValue is null && convertedValue is null) isDifferent = false;
        else if (propertyType == typeof(string) || propertyType.IsEnum)
            isDifferent = !string.Equals(currentValue?.ToString()?.Trim(), convertedValue?.ToString()?.Trim(), StringComparison.OrdinalIgnoreCase);
        else
            isDifferent = !Equals(currentValue, convertedValue);

        if (!isDifferent)
            return false;

        propertyInfo.SetValue(ticket, convertedValue);
        return true;
    }
    catch { return false; }
    }

    /// <summary>
    /// Converte string para o tipo de destino da propriedade do ticket.
    ///
    /// Suporta enums, tipos primitivos e <c>Nullable&lt;T&gt;</c>. Se a conversao
    /// falhar e o destino for string, preserva o valor original; caso contrario,
    /// propaga a excecao para o chamador decidir o tratamento.
    /// </summary>
    private static object? ConvertIfPossible(Type targetType, string rawValue)
    {
        try
        {
            var isNullable = targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(Nullable<>);
            var underlyingType = isNullable ? Nullable.GetUnderlyingType(targetType) : null;

            if (isNullable && string.IsNullOrWhiteSpace(rawValue))
                return null;
            if ((underlyingType ?? targetType).IsEnum)
            {
                var enumType = (underlyingType ?? targetType);
                return Enum.Parse(enumType, rawValue, ignoreCase: true);
            }
            var realType = underlyingType ?? targetType;
            return System.Convert.ChangeType(rawValue, realType);
        }
        catch
        {
            if (targetType == typeof(string) || (Nullable.GetUnderlyingType(targetType) == typeof(string)))
                return rawValue;
            throw;
        }
    }

    /// <summary>
    /// Descarta o objeto quando ele implementa <see cref="IDisposable"/>.
    ///
    /// Helper para limpeza segura de instancias criadas por reflection.
    /// </summary>
    private static void DisposeIfPossible(object? obj)
    {
        (obj as IDisposable)?.Dispose();
    }

    /// <summary>
    /// Cria uma instancia de <c>LocalPrintServer</c> com acesso desejado.
    ///
    /// Tenta usar o construtor que recebe <c>PrintSystemDesiredAccess</c>.
    /// Se o tipo/acesso nao estiver disponivel, faz fallback para construtor
    /// padrao para maximizar compatibilidade entre ambientes.
    /// </summary>
    private static object? CreateLocalPrintServer(Type printServerType, string requestedAccessName)
    {
        try
        {
            var desiredAccessType = Type.GetType("System.Printing.PrintSystemDesiredAccess, System.Printing", throwOnError: false);
            if (desiredAccessType is null)
                return Activator.CreateInstance(printServerType);

            var desiredAccess = Enum.Parse(desiredAccessType, requestedAccessName, ignoreCase: true);
            var ctor = printServerType.GetConstructor(new[] { desiredAccessType });
            if (ctor is null)
                return Activator.CreateInstance(printServerType);

            return ctor.Invoke(new[] { desiredAccess });
        }
        catch
        {
            return Activator.CreateInstance(printServerType);
        }
    }

    /// <summary>
    /// Obtem a fila de impressao por nome, com fallback entre estrategias.
    ///
    /// Prioriza criacao direta de <c>PrintQueue</c> com acesso explicito.
    /// Quando isso nao e possivel, tenta <c>GetPrintQueue</c> com lista de
    /// propriedades e, por fim, a sobrecarga simples por nome.
    /// </summary>
    private static object? GetPrintQueue(
        Type printServerType,
        object? printServer,
        string queueName,
        string requiredTicketPropertyName,
        string requestedAccessName)
    {
        if (printServer is null)
            return null;

        var desiredAccessType = Type.GetType("System.Printing.PrintSystemDesiredAccess, System.Printing", throwOnError: false);
        if (desiredAccessType is not null)
        {
            var desiredAccess = Enum.Parse(desiredAccessType, requestedAccessName, ignoreCase: true);

            var printQueueType = Type.GetType("System.Printing.PrintQueue, System.Printing", throwOnError: false);
            var basePrintServerType = Type.GetType("System.Printing.PrintServer, System.Printing", throwOnError: false);
            if (printQueueType is not null && basePrintServerType is not null && basePrintServerType.IsInstanceOfType(printServer))
            {
                var ctor = printQueueType.GetConstructor(new[] { basePrintServerType, typeof(string), desiredAccessType });
                if (ctor is not null)
                    return ctor.Invoke(new object[] { printServer, queueName, desiredAccess });
            }
        }

        var getQueueWithProperties = printServerType.GetMethod("GetPrintQueue", new[] { typeof(string), typeof(string[]) });
        if (getQueueWithProperties is not null)
        {
            var properties = new[] { requiredTicketPropertyName };
            return getQueueWithProperties.Invoke(printServer, new object[] { queueName, properties });
        }

        var getQueueDefault = printServerType.GetMethod("GetPrintQueue", new[] { typeof(string) });
        return getQueueDefault?.Invoke(printServer, new object[] { queueName });
    }

    /// <summary>
    /// Retorna a mensagem da excecao mais interna da cadeia.
    ///
    /// Util para expor causa raiz de erros de reflection/invocacao sem trazer
    /// o stack trace completo para o resultado funcional.
    /// </summary>
    private static string GetInnermostMessage(Exception ex)
    {
        var current = ex;
        while (current.InnerException is not null)
            current = current.InnerException;

        return current.Message;
    }

        
}
