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
        private const string FIXED_DUPLEX = "TwoSidedLongEdge";
        private const string FIXED_COLOR = "Monochrome";
        private const string FIXED_ORIENTATION = "Portrait";

        public static PrintTicketEnforcementResult EnforceDefaultTicketPolicy(
            IPrintTicketService service,
            string queueName)
        {
            var info = service.GetDefaultTicketInfo(queueName);
            if (!info.Available)
            {
                return new PrintTicketEnforcementResult(false, "Não foi possível ler PrintTicket: " + info.Details, false);
            }

            var changes = new Dictionary<string, string?>();
            bool requiresUpdate = false;

            info.Attributes.TryGetValue("Duplexing", out var curDuplex);
            if (!string.Equals(curDuplex, FIXED_DUPLEX, StringComparison.OrdinalIgnoreCase))
            {
                changes["Duplexing"] = FIXED_DUPLEX;
                requiresUpdate = true;
            }

            info.Attributes.TryGetValue("OutputColor", out var curColor);
            if (!string.Equals(curColor, FIXED_COLOR, StringComparison.OrdinalIgnoreCase))
            {
                changes["OutputColor"] = FIXED_COLOR;
                requiresUpdate = true;
            }

            info.Attributes.TryGetValue("PageOrientation", out var curOrientation);
            if (!string.Equals(curOrientation, FIXED_ORIENTATION, StringComparison.OrdinalIgnoreCase))
            {
                changes["PageOrientation"] = FIXED_ORIENTATION;
                requiresUpdate = true;
            }

            if (!requiresUpdate)
            {
                return new PrintTicketEnforcementResult(true, "Já em conformidade com a política.", false);
            }

            var request = new PrintTicketUpdateRequest(
                changes.ContainsKey("Duplexing") ? changes["Duplexing"] : null,
                changes.ContainsKey("OutputColor") ? changes["OutputColor"] : null,
                changes.ContainsKey("PageOrientation") ? changes["PageOrientation"] : null);

            var update = service.UpdateDefaultTicket(queueName, request);

            if (update.Applied)
            {
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("Enforcement realizado com sucesso.");
                sb.AppendLine($"  Duplexing: {curDuplex} → {FIXED_DUPLEX}");
                sb.AppendLine($"  OutputColor: {curColor} → {FIXED_COLOR}");
                sb.AppendLine($"  PageOrientation: {curOrientation} → {FIXED_ORIENTATION}");
                return new PrintTicketEnforcementResult(true, sb.ToString(), true);
            }
            else
            {
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
}


