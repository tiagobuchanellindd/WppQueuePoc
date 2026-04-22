# Monitoramento e Enforcement Automático de Propriedades de Impressora via Win32

**Referência rápida para estudo:**

- As classes-exemplo `WinPrinterChangesMonitor.cs` e `PrinterExternalDeletion.cs`, nesta pasta, apresentam implementações reais de monitoramento de impressoras (via Win32 APIs, subscrição de eventos, manipulação robusta de notificações e integração com estratégias de sincronização). Podem ajudar a consulta, inspiração ou troubleshooting dos fluxos de monitoração, sincronismo e tratamento de eventos do Windows.
- Para detalhes oficiais e exemplos sobre monotoração de mudança de impressora via Windows consulte a documentação Microsoft:
  - [FindFirstPrinterChangeNotification (docs)](https://learn.microsoft.com/en-us/windows/win32/printdocs/findfirstprinterchangenotification)

---

## 1. Objetivo
Desenvolver um novo componente que monitora alterações nas propriedades de impressoras do Windows (por exemplo, Duplex, Orientação, Cor) e restaura automaticamente para valores definidos por política sempre que uma alteração for detectada.

- Não reaproveitar classes existentes, mas manter o alinhamento técnico (event-based, Win32 APIs).
- Permitir controle pelo usuário via MainWindow (botão para ativar/desativar, status/output).

## 2. Etapas e Detalhes Técnicos

### 2.1. Especificação e Coleta de Requisitos
- Identificar propriedades que serão protegidas (Duplex, Orientação, Cor etc).
- Definir política de valores aceitos/forçados por impressora.

### 2.2. Projeto do Componente
- Nova classe, sugerido: `PrinterPolicyEnforcer`.
- Métodos principais: `StartMonitoring`, `StopMonitoring`.
- Estado: política carregada/padrão, impressora(s) alvo.
- Integração com UI: feedback de status pelo OutputTextBox em MainWindow.
- Isolamento, sem dependência de código legado.

### 2.3. Subscrever e Monitorar Eventos
- Utilizar `FindFirstPrinterChangeNotification` (PRINTER_CHANGE_SET_PRINTER...).
- Abrir handle da impressora, iniciar thread/task de background.
- Loop: aguarda eventos (WaitForSingleObject/FindNextPrinterChangeNotification), processa.
- Encerrar monitoramento com recurso seguro (liberar handles/threads).

### 2.4. Detecção e Enforcement
- Ao receber evento relevante:
  - Consultar estado atual do PrintTicket da fila (IPrintTicketService já existente).
  - Comparar atributos críticos x política definida.
  - Se diferente, aplicar correção via fluxo de update de PrintTicket (já modelado).
  - Log/output resumido na interface.
- Evitar loop recursivo ou flood (debounce, controle de origem do update).

### 2.5. Integração na UI
- Novo botão em MainWindow: "Ativar monitoramento de políticas de impressora".
- Mostrar status/output na OutputTextBox.
- Permitir feedback de sucesso/falha para o usuário.

### 2.6. Testes
- Alteração manual: garantir rollback automático.
- Troca de propriedades não cobertas: não agir.
- Permissão insuficiente: logar falha, não quebrar rotina.
- Parar/iniciar monitoramento repetidas vezes, sem vazamento de recursos.

### 2.7. Considerações Avançadas
- Suporte futuro para múltiplas impressoras/políticas dinâmicas.
- Compatibilidade e fallback para drivers antigos.
- Extensível para incluir notificações/history na UI (não obrigatório no MVP).

## 3. Checklist de Implementação

- [ ] Mapear e detalhar políticas de propriedades.
- [ ] Projetar classe "enforcer" com interface clara e eventos de output.
- [ ] Codificar subscrição Win32 e loop de eventos (thread segura, cancelável).
- [ ] Integrar com rotina de leitura/atualização de PrintTicket (serviço já pronto).
- [ ] Implementar debounce/proteção contra loops recursivos.
- [ ] Ajustar UI: botão de ativar/parar e display de eventos/status.
- [ ] Testar todos os fluxos (positivo, negativo, abandono early, sem permissão).
- [ ] Documentar limitações e extensões desejáveis.

## 4. Observações Finais
- Não modificar ou impactar as classes de monitoramento existentes no produto.
- O componente deve ser autoexplicativo e fácil de remover/testar.
- Qualquer mecanismo de fallback (polling, try/catch robustness) deve ser documentado na própria classe.

---

**Este plano deve ser salvo em `docs/feature-printer-policy-enforcer.md` e consultado durante todas as fases do desenvolvimento.**
