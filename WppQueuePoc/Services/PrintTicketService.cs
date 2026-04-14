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
        var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var localPrintServerType = Type.GetType("System.Printing.LocalPrintServer, System.Printing", throwOnError: false);
        if (localPrintServerType is null)
        {
            return new PrintTicketInfoResult(
                queueName,
                Available: false,
                Details: "System.Printing assembly is not available in this runtime.",
                Attributes: attributes);
        }
        object? localPrintServer = null;
        object? printQueue = null;
        try
        {
            localPrintServer = CreateLocalPrintServer(localPrintServerType, "AdministrateServer");
            if (localPrintServer is null)
            {
                return new PrintTicketInfoResult(queueName, false, "Unable to create LocalPrintServer.", attributes);
            }
            printQueue = GetPrintQueue(localPrintServerType, localPrintServer, queueName, "DefaultPrintTicket", "AdministratePrinter");
            if (printQueue is null)
            {
                return new PrintTicketInfoResult(queueName, false, $"Queue '{queueName}' not found.", attributes);
            }
            var queueType = printQueue.GetType();
            var defaultTicketProperty = queueType.GetProperty("DefaultPrintTicket");
            var defaultTicket = defaultTicketProperty?.GetValue(printQueue);
            if (defaultTicket is null)
            {
                return new PrintTicketInfoResult(queueName, false, "DefaultPrintTicket is not available.", attributes);
            }
            ReadTicketAttribute(defaultTicket, "OutputColor", attributes);
            ReadTicketAttribute(defaultTicket, "PageMediaSize", attributes);
            ReadTicketAttribute(defaultTicket, "PageOrientation", attributes);
            ReadTicketAttribute(defaultTicket, "InputBin", attributes);
            ReadTicketAttribute(defaultTicket, "Duplexing", attributes);
            ReadTicketAttribute(defaultTicket, "CopyCount", attributes);
            ReadTicketAttribute(defaultTicket, "Collation", attributes);
            ReadTicketAttribute(defaultTicket, "Stapling", attributes);
            return new PrintTicketInfoResult(
                queueName,
                Available: true,
                Details: "Default print ticket information captured.",
                Attributes: attributes);
        }
        catch (Exception ex)
        {
            return new PrintTicketInfoResult(
                queueName,
                Available: false,
                Details: $"Failed to read default print ticket information: {ex.Message}",
                Attributes: attributes);
        }
        finally
        {
            DisposeIfPossible(printQueue);
            DisposeIfPossible(localPrintServer);
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
        var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var localPrintServerType = Type.GetType("System.Printing.LocalPrintServer, System.Printing", throwOnError: false);
        if (localPrintServerType is null)
        {
            return new PrintTicketInfoResult(
                queueName,
                Available: false,
                Details: "System.Printing assembly is not available in this runtime.",
                Attributes: attributes);
        }
        object? localPrintServer = null;
        object? printQueue = null;
        try
        {
            localPrintServer = CreateLocalPrintServer(localPrintServerType, "AdministrateServer");
            if (localPrintServer is null)
            {
                return new PrintTicketInfoResult(queueName, false, "Unable to create LocalPrintServer.", attributes);
            }
            printQueue = GetPrintQueue(localPrintServerType, localPrintServer, queueName, "UserPrintTicket", "AdministratePrinter");
            if (printQueue is null)
            {
                return new PrintTicketInfoResult(queueName, false, $"Queue '{queueName}' not found.", attributes);
            }
            var queueType = printQueue.GetType();
            var userTicketProperty = queueType.GetProperty("UserPrintTicket");
            var userTicket = userTicketProperty?.GetValue(printQueue);
            if (userTicket is null)
            {
                return new PrintTicketInfoResult(queueName, false, "UserPrintTicket is not available.", attributes);
            }
            ReadTicketAttribute(userTicket, "OutputColor", attributes);
            ReadTicketAttribute(userTicket, "PageMediaSize", attributes);
            ReadTicketAttribute(userTicket, "PageOrientation", attributes);
            ReadTicketAttribute(userTicket, "InputBin", attributes);
            ReadTicketAttribute(userTicket, "Duplexing", attributes);
            ReadTicketAttribute(userTicket, "CopyCount", attributes);
            ReadTicketAttribute(userTicket, "Collation", attributes);
            ReadTicketAttribute(userTicket, "Stapling", attributes);
            return new PrintTicketInfoResult(
                queueName,
                Available: true,
                Details: "User print ticket information captured.",
                Attributes: attributes);
        }
        catch (Exception ex)
        {
            return new PrintTicketInfoResult(
                queueName,
                Available: false,
                Details: $"Failed to read user print ticket information: {ex.Message}",
                Attributes: attributes);
        }
        finally
        {
            DisposeIfPossible(printQueue);
            DisposeIfPossible(localPrintServer);
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
    string ticketTypeProperty,
    string scope)
    {
        var requested = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        void TryAdd(string key, string? value)
        {
            if (!string.IsNullOrWhiteSpace(value))
                requested[key] = value.Trim();
        }
        TryAdd("Duplexing", request.Duplexing);
        TryAdd("OutputColor", request.OutputColor);
        var localPrintServerType = Type.GetType("System.Printing.LocalPrintServer, System.Printing", throwOnError: false);
        if (localPrintServerType is null)
        {
            return new PrintTicketUpdateResult(
                queueName,
                scope,
                false,
                "System.Printing not available.",
                requested,
                new Dictionary<string, string>());
        }
        object? localPrintServer = null;
        object? printQueue = null;
        try
        {
            localPrintServer = CreateLocalPrintServer(localPrintServerType, "AdministrateServer");
            printQueue = GetPrintQueue(localPrintServerType, localPrintServer, queueName, ticketTypeProperty, "AdministratePrinter");
            if (printQueue is null)
                return new PrintTicketUpdateResult(
                    queueName,
                    scope,
                    false,
                    $"Queue '{queueName}' not found.",
                    requested,
                    new Dictionary<string, string>());
            var queueType = printQueue.GetType();
            var ticketProperty = queueType.GetProperty(ticketTypeProperty);
            if(ticketProperty == null)
                return new PrintTicketUpdateResult(
                    queueName,
                    scope,
                    false,
                    $"Print ticket property '{ticketTypeProperty}' not found.",
                    requested,
                    new Dictionary<string, string>());
            var ticket = ticketProperty.GetValue(printQueue);
            if (ticket is null)
                return new PrintTicketUpdateResult(
                    queueName,
                    scope,
                    false,
                    $"Print ticket '{ticketTypeProperty}' not found (might be unsupported by this driver or printer type).",
                    requested,
                    new Dictionary<string, string>());
            // Aplica os atributos
            bool changed = false;
            changed |= WriteTicketAttribute(ticket, "Duplexing", request.Duplexing);
            changed |= WriteTicketAttribute(ticket, "OutputColor", request.OutputColor);
            changed |= WriteTicketAttribute(ticket, "PageOrientation", request.PageOrientation); // Now applies PageOrientation updates too!
            string? commitError = null;
            if (changed)
            {
                try
                {
                    // Always set the ticket back per MS documentation!
                    ticketProperty.SetValue(printQueue, ticket);
                }
                catch(Exception setEx)
                {
                    var error = GetInnermostMessage(setEx);
                    commitError = $"Failed to set updated ticket back on '{ticketTypeProperty}': {error}";
                }

                try
                {
                    // Only commit after set
                    var commit = queueType.GetMethod("Commit");
                    commit?.Invoke(printQueue, null);
                }
                catch(Exception commitEx)
                {
                    // Capture possible commit failure
                    var error = GetInnermostMessage(commitEx);
                    commitError = $"Commit failed: {error}";
                }
            }
            // Lê os valores efetivos após possível commit/aplicação
            var applied = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            ReadTicketAttribute(ticket, "Duplexing", applied);
            ReadTicketAttribute(ticket, "OutputColor", applied);
            ReadTicketAttribute(ticket, "PageOrientation", applied);
            // Result message
            string resultMsg = changed
                ? (commitError == null ? "Ticket updated successfully." : $"Ticket update attempted. {commitError}")
                : "No changes needed: the provided values matched current settings.";

            return new PrintTicketUpdateResult(
                queueName,
                scope,
                changed,
                resultMsg,
                requested,
                applied);
        }
        catch (Exception ex)
        {
            var error = GetInnermostMessage(ex);
            return new PrintTicketUpdateResult(
                queueName,
                scope,
                false,
                $"Exception during PrintTicket update: {error}. This may occur if the process is not elevated and/or the account does not have 'Manage Printers' permission on this queue.",
                requested,
                new Dictionary<string, string>());
        }
        finally
        {
            DisposeIfPossible(printQueue);
            DisposeIfPossible(localPrintServer);
        }
    }

    /// <summary>
    /// Le uma propriedade do ticket por reflection e grava no dicionario de saida.
    ///
    /// A leitura e best-effort: se a propriedade nao existir ou ocorrer excecao,
    /// o metodo ignora a falha para nao interromper a coleta dos demais atributos.
    /// </summary>
    private static void ReadTicketAttribute(object? ticket, string attrName, IDictionary<string, string> output)
    {
        try
        {
            if (ticket == null) return;
            var type = ticket.GetType();
            var prop = type.GetProperty(attrName);
            if (prop != null)
            {
                var value = prop.GetValue(ticket);
                output[attrName] = value?.ToString() ?? "";
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
    private static bool WriteTicketAttribute(object? ticket, string attrName, string? value)
{
    try
    {
        if (ticket == null || string.IsNullOrWhiteSpace(value))
            return false;
        var type = ticket.GetType();
        var prop = type.GetProperty(attrName);
        if (prop == null || !prop.CanWrite) return false;

        // Obtain the current value
        var currentValue = prop.GetValue(ticket);
        var targetType = prop.PropertyType;
        object? realValue = ConvertIfPossible(targetType, value);

        // Compare by string representation for enums/strings, deep equals otherwise
        bool isDifferent;
        if (currentValue is null && realValue is not null) isDifferent = true;
        else if (currentValue is not null && realValue is null) isDifferent = true;
        else if (currentValue is null && realValue is null) isDifferent = false;
        else if (targetType == typeof(string) || targetType.IsEnum)
            isDifferent = !string.Equals(currentValue?.ToString()?.Trim(), realValue?.ToString()?.Trim(), StringComparison.OrdinalIgnoreCase);
        else
            isDifferent = !Equals(currentValue, realValue);

        if (!isDifferent)
            return false;

        prop.SetValue(ticket, realValue);
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
    private static object? ConvertIfPossible(Type targetType, string value)
    {
        try
        {
            // Handle Nullable<T> for enums and value types
            var isNullable = targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(Nullable<>);
            var underlyingType = isNullable ? Nullable.GetUnderlyingType(targetType) : null;

            if (isNullable && string.IsNullOrWhiteSpace(value))
                return null;
            if ((underlyingType ?? targetType).IsEnum)
            {
                var enumType = (underlyingType ?? targetType);
                return Enum.Parse(enumType, value, ignoreCase: true);
            }
            // Handle common value type conversion, including underlying types for nullables
            var realType = underlyingType ?? targetType;
            return System.Convert.ChangeType(value, realType);
        }
        catch
        {
            // Only fallback for string destination; otherwise throw
            if (targetType == typeof(string) || (Nullable.GetUnderlyingType(targetType) == typeof(string)))
                return value;
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
    private static object? CreateLocalPrintServer(Type localPrintServerType, string desiredAccessName)
    {
        try
        {
            var desiredAccessType = Type.GetType("System.Printing.PrintSystemDesiredAccess, System.Printing", throwOnError: false);
            if (desiredAccessType is null)
                return Activator.CreateInstance(localPrintServerType);

            var desiredAccess = Enum.Parse(desiredAccessType, desiredAccessName, ignoreCase: true);
            var ctor = localPrintServerType.GetConstructor(new[] { desiredAccessType });
            if (ctor is null)
                return Activator.CreateInstance(localPrintServerType);

            return ctor.Invoke(new[] { desiredAccess });
        }
        catch
        {
            return Activator.CreateInstance(localPrintServerType);
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
        Type localPrintServerType,
        object? localPrintServer,
        string queueName,
        string ticketProperty,
        string desiredAccessName)
    {
        if (localPrintServer is null)
            return null;

        var desiredAccessType = Type.GetType("System.Printing.PrintSystemDesiredAccess, System.Printing", throwOnError: false);
        if (desiredAccessType is not null)
        {
            var desiredAccess = Enum.Parse(desiredAccessType, desiredAccessName, ignoreCase: true);

            var printQueueType = Type.GetType("System.Printing.PrintQueue, System.Printing", throwOnError: false);
            var printServerType = Type.GetType("System.Printing.PrintServer, System.Printing", throwOnError: false);
            if (printQueueType is not null && printServerType is not null && printServerType.IsInstanceOfType(localPrintServer))
            {
                var ctor = printQueueType.GetConstructor(new[] { printServerType, typeof(string), desiredAccessType });
                if (ctor is not null)
                    return ctor.Invoke(new object[] { localPrintServer, queueName, desiredAccess });
            }
        }

        var getQueueWithProperties = localPrintServerType.GetMethod("GetPrintQueue", new[] { typeof(string), typeof(string[]) });
        if (getQueueWithProperties is not null)
        {
            var properties = new[] { ticketProperty };
            return getQueueWithProperties.Invoke(localPrintServer, new object[] { queueName, properties });
        }

        var getQueueDefault = localPrintServerType.GetMethod("GetPrintQueue", new[] { typeof(string) });
        return getQueueDefault?.Invoke(localPrintServer, new object[] { queueName });
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
