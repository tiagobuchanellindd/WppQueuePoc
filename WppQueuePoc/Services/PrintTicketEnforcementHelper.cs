using WppQueuePoc.Models;
using WppQueuePoc.Services;
using System;
using System.Collections.Generic;
using WppQueuePoc.Abstractions;

namespace WppQueuePoc.Services
{
    /// <summary>
    /// Helper para comparar e aplicar enforcement de política de PrintTicket.
    /// Mantém lógica isolada, sem alterar contratos existentes.
    /// </summary>
    public static class PrintTicketEnforcementHelper
    {
        public static PrintTicketEnforcementResult EnforceDefaultTicketPolicy(
            IPrintTicketService service,
            string queueName,
            PrinterPolicyEnforcer.Policy policy)
        {
            // 1. Lê ticket atual
            var info = service.GetDefaultTicketInfo(queueName);
            if (!info.Available)
            {
                return new PrintTicketEnforcementResult(false, "Não foi possível ler PrintTicket atual: " + info.Details, false);
            }
            var changes = new Dictionary<string, string?>();
            bool requiresUpdate = false;
            if (policy.EnforceDuplex && policy.RequiredDuplexValue != null)
            {
                info.Attributes.TryGetValue("Duplexing", out var curValue);
                if (!string.Equals(curValue, policy.RequiredDuplexValue, StringComparison.OrdinalIgnoreCase))
                {
                    changes["Duplexing"] = policy.RequiredDuplexValue;
                    requiresUpdate = true;
                }
            }
            if (policy.EnforceColor && policy.RequiredColorValue != null)
            {
                info.Attributes.TryGetValue("OutputColor", out var curValue);
                if (!string.Equals(curValue, policy.RequiredColorValue, StringComparison.OrdinalIgnoreCase))
                {
                   changes["OutputColor"] = policy.RequiredColorValue;
                   requiresUpdate = true;
                }
            }
            if (policy.EnforceOrientation && policy.RequiredOrientationValue != null)
            {
                info.Attributes.TryGetValue("PageOrientation", out var curValue);
                if (!string.Equals(curValue, policy.RequiredOrientationValue, StringComparison.OrdinalIgnoreCase))
                {
                   changes["PageOrientation"] = policy.RequiredOrientationValue;
                   requiresUpdate = true;
                }
            }

            if (!requiresUpdate)
            {
                return new PrintTicketEnforcementResult(true, "Já em conformidade com a política.", false);
            }

            var request = new PrintTicketUpdateRequest(
                changes.ContainsKey("Duplexing") ? changes["Duplexing"] : null,
                changes.ContainsKey("OutputColor") ? changes["OutputColor"] : null,
                changes.ContainsKey("PageOrientation") ? changes["PageOrientation"] : null
            );
            var update = service.UpdateDefaultTicket(queueName, request);

            if (update.Applied)
                return new PrintTicketEnforcementResult(true, "Enforcement realizado com sucesso.", true);
            else
                return new PrintTicketEnforcementResult(false, "Falha ao aplicar enforcement: " + update.Details, true);
        }
    }

    /// <summary>
    /// Resultado estruturado do enforcement.
    /// </summary>
    public sealed record PrintTicketEnforcementResult(
        bool Success,
        string Details,
        bool AttemptedChange);
}
