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
            string commitError = null;
            if (changed)
            {
                try
                {
                    // Always set the ticket back per MS documentation!
                    ticketProperty.SetValue(printQueue, ticket);
                }
                catch(Exception setEx)
                {
                    commitError = $"Failed to set updated ticket back on '{ticketTypeProperty}': {setEx.Message}";
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
                    commitError = $"Commit failed: {commitEx.Message}";
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
            return new PrintTicketUpdateResult(
                queueName,
                scope,
                false,
                $"Exception during PrintTicket update: {ex.Message}. This may occur if the printer driver does not support programmatic changes to the default or user print ticket.",
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
    private static void DisposeIfPossible(object? obj)
    {
        (obj as IDisposable)?.Dispose();
    }

        
}