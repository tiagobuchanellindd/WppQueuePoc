using WppQueuePoc.Models;
using WppQueuePoc.Services;
using System;
using System.Collections.Generic;
using WppQueuePoc.Abstractions;

namespace WppQueuePoc.Services
{
    /// <summary>
    /// Helper estático que compara configurações atuais do PrintTicket com políticas
    /// e aplica correções quando necessário.
    ///
    /// Esta classe implementa o núcleo do padrão "enforcement": compara os valores
    /// atuais do DefaultPrintTicket (lidos via IPrintTicketService) com os valores
    /// requeridos pela política e, se houver divergência, chama a atualização
    /// para restaurar os valores corretos.
    ///
    /// O fluxo de enforcement:
    /// 1. Le o DefaultPrintTicket atual da fila via GetDefaultTicketInfo
    /// 2. Para cada dimensão habilitada na política (Duplex, Color, Orientation),
    ///    compara o valor atual com o requerido
    /// 3. Se houver diferença, prepara request de UpdateDefaultTicket
    /// 4. Executa a atualização e retorna resultado estruturado
    ///
    /// Esta abordagem isolada (fora do PrinterPolicyEnforcer) permite que o helper
    /// seja 测试ado unitariamente e reutilizado em outros contextos.
    /// </summary>
    public static class PrintTicketEnforcementHelper
    {
        /// <summary>
        /// Executa o enforcement de políticas no DefaultPrintTicket de uma fila.
        ///
        /// Este método é o punto de entrada do helper. Ele:
        /// -Lê o ticket atual (via service.GetDefaultTicketInfo)
        /// -Compara cada propriedade habilitada na política
        /// -Se necessario, envia atualização (via service.UpdateDefaultTicket)
        /// -Retorna um PrintTicketEnforcementResult com detalhes do resultado
        ///
        /// A lógica de comparação é case-insensitive (OrdinalIgnoreCase) para
        /// tolerar variações de casing nos valores dos drivers.
        /// </summary>
        /// <param name="service">Instancia de IPrintTicketService para operas na fila.</param>
        /// <param name="queueName">Nome da fila de impressao.</param>
        /// <param name="policy">Politica com parametros a enforcing.</param>
        /// <returns>PrintTicketEnforcementResult com sucesso, detalhes e se houve mudanca.</returns>
        public static PrintTicketEnforcementResult EnforceDefaultTicketPolicy(
        IPrintTicketService service,
        string queueName,
        PrinterPolicyEnforcer.Policy policy)
        {
            var info = service.GetDefaultTicketInfo(queueName);
            if (!info.Available)
            {
                return new PrintTicketEnforcementResult(false, "Não foi possível ler PrintTicket: " + info.Details, false);
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
                changes.ContainsKey("PageOrientation") ? changes["PageOrientation"] : null);

            var update = service.UpdateDefaultTicket(queueName, request);

            if (update.Applied)
            {
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("Enforcement realizado com sucesso.");
                if (changes.ContainsKey("Duplexing"))
                    sb.AppendLine($"  Duplexing: {info.Attributes["Duplexing"]} → {policy.RequiredDuplexValue}");
                if (changes.ContainsKey("OutputColor"))
                    sb.AppendLine($"  OutputColor: {info.Attributes["OutputColor"]} → {policy.RequiredColorValue}");
                if (changes.ContainsKey("PageOrientation"))
                    sb.AppendLine($"  PageOrientation: {info.Attributes["PageOrientation"]} → {policy.RequiredOrientationValue}");
                return new PrintTicketEnforcementResult(true, sb.ToString(), true);
            }
            else
            {
                return new PrintTicketEnforcementResult(false, "Falha ao aplicar enforcement: " + update.Details, true);
            }
        }

        /// <summary>
        /// Resultado estruturado da operacao de enforcement.
        ///
        /// Agrupa success (bool global), details (mensagem descritiva) e attemptedChange
        /// (se houve tentativa de mudanca). Isso permite ao chamador distinguir
        /// entre "ja conforme" (= sem mudanca, mas sucesso) e "tentou mudar" (= aplicou request).
        /// </summary>
        public sealed record PrintTicketEnforcementResult(
            /// <summary>
            /// Indica se a operacao overall foi bem-sucedida.
            /// </summary>
            bool Success,
            /// <summary>
            /// Mensagem descritiva com detalhes do resultado.
            /// </summary>
            string Details,
            /// <summary>
            /// Indica se houve tentativa real de mudanca (vs. ja conforme).
            /// </summary>
            bool AttemptedChange);

    }
}


