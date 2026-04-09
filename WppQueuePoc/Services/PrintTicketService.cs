using WppQueuePoc.Abstractions;
using WppQueuePoc.Models;
using System;
using System.Collections.Generic;
namespace WppQueuePoc.Services;

public sealed partial class PrintTicketService : IPrintTicketService
{
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
            localPrintServer = Activator.CreateInstance(localPrintServerType);
            if (localPrintServer is null)
            {
                return new PrintTicketInfoResult(queueName, false, "Unable to create LocalPrintServer.", attributes);
            }
            var getPrintQueue = localPrintServerType.GetMethod("GetPrintQueue", new[] { typeof(string) });
            if (getPrintQueue is null)
            {
                return new PrintTicketInfoResult(queueName, false, "GetPrintQueue method not found.", attributes);
            }
            printQueue = getPrintQueue.Invoke(localPrintServer, new object[] { queueName });
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
            localPrintServer = Activator.CreateInstance(localPrintServerType);
            if (localPrintServer is null)
            {
                return new PrintTicketInfoResult(queueName, false, "Unable to create LocalPrintServer.", attributes);
            }
            var getPrintQueue = localPrintServerType.GetMethod("GetPrintQueue", new[] { typeof(string) });
            if (getPrintQueue is null)
            {
                return new PrintTicketInfoResult(queueName, false, "GetPrintQueue method not found.", attributes);
            }
            printQueue = getPrintQueue.Invoke(localPrintServer, new object[] { queueName });
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
    public PrintTicketUpdateResult UpdateDefaultTicket(string queueName, PrintTicketUpdateRequest request)
    => UpdatePrintTicketInternal(queueName, request, "DefaultPrintTicket", "Default");
    public PrintTicketUpdateResult UpdateUserTicket(string queueName, PrintTicketUpdateRequest request)
        => UpdatePrintTicketInternal(queueName, request, "UserPrintTicket", "User");

    // Core updater for both default/user ticket
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
            localPrintServer = Activator.CreateInstance(localPrintServerType);
            var getPrintQueue = localPrintServerType.GetMethod("GetPrintQueue", new[] { typeof(string) });
            printQueue = getPrintQueue?.Invoke(localPrintServer, new object[] { queueName });
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
            var ticket = ticketProperty?.GetValue(printQueue);
            if (ticket is null)
                return new PrintTicketUpdateResult(
                    queueName,
                    scope,
                    false,
                    $"Print ticket '{ticketTypeProperty}' not found.",
                    requested,
                    new Dictionary<string, string>());
            // Aplica os atributos
            bool changed = false;
            changed |= WriteTicketAttribute(ticket, "Duplexing", request.Duplexing);
            changed |= WriteTicketAttribute(ticket, "OutputColor", request.OutputColor);
            if (changed)
            {
                // Salva no spooler
                var commit = queueType.GetMethod("Commit");
                commit?.Invoke(printQueue, null);
            }
            // Lê os valores efetivos após possível commit/aplicação
            var applied = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            ReadTicketAttribute(ticket, "Duplexing", applied);
            ReadTicketAttribute(ticket, "OutputColor", applied);
            ReadTicketAttribute(ticket, "PageOrientation", applied);
            return new PrintTicketUpdateResult(
                queueName,
                scope,
                changed,
                changed ? "Ticket updated successfully." : "No changes applied (all values same as before).",
                requested,
                applied);
        }
        catch (Exception ex)
        {
            return new PrintTicketUpdateResult(
                queueName,
                scope,
                false,
                $"Exception: {ex.Message}",
                requested,
                new Dictionary<string, string>());
        }
        finally
        {
            DisposeIfPossible(printQueue);
            DisposeIfPossible(localPrintServer);
        }
    }

    // Reflection utility for ticket reading
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
    // Reflection utility for ticket writing
    private static bool WriteTicketAttribute(object? ticket, string attrName, string? value)
    {
        try
        {
            if (ticket == null || string.IsNullOrWhiteSpace(value))
                return false;
            var type = ticket.GetType();
            var prop = type.GetProperty(attrName);
            if (prop == null || !prop.CanWrite) return false;
            // Attempt to convert value if needed (string parsing)
            var targetType = prop.PropertyType;
            object? realValue = ConvertIfPossible(targetType, value);
            prop.SetValue(ticket, realValue);
            return true;
        }
        catch { return false; }
    }
    private static object? ConvertIfPossible(Type targetType, string value)
    {
        try
        {
            if (targetType.IsEnum)
                return Enum.Parse(targetType, value, ignoreCase: true);
            return System.Convert.ChangeType(value, targetType);
        }
        catch
        {
            return value; // fallback: set string
        }
    }
    private static void DisposeIfPossible(object? obj)
    {
        (obj as IDisposable)?.Dispose();
    }

        
}