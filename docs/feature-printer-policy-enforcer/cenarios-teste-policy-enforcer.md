# Cenários de Teste — PrinterPolicyEnforcer

## Cenário 1: Enforcement automático após alteração manual

**Objetivo:** Verificar que uma alteração manual em Duplex, Cor ou Orientação na fila é revertida automaticamente em conformidade com a policy.

### Passos
1. Inicie o PoolPrinting/WppQueuePoc normalmente.
2. Na interface, clique em "Start Monitoring" (ou botão equivalente de ativação do monitor).
3. No Windows, vá em "Dispositivos e Impressoras", localize "Microsoft Print to PDF" (ou fila monitorada pelo MVP).
4. Clique com o botão direito > Preferências de Impressão.
5. Altere o valor de Duplex, Cor ou Orientação para um valor diferente da policy atual:
    - Exemplo: Policy exige "Monochrome" — altere para "Colorido".
    - Policy exige "Portrait" — altere para "Paisagem".
6. Clique em "OK" para salvar a alteração.
7. Volte para o software PoolPrinting e observe:
    - Deve aparecer um log indicando "Mudança capturada na impressora..."
    - Em seguida, um log de "Enforcement realizado com sucesso." ou similar.
    - Se o enforcement funcionar, a fila volta à configuração da policy após alguns segundos.
8. Confirme voltando à janela de preferências da fila: o valor voltou à configuração da policy.

---
## Cenário 2: Debounce — Mudanças próximas são ignoradas

**Objetivo:** Validar que o mecanismo debounce impede enforcement repetido em sequência.

### Passos
1. Repita os passos 1–5 do cenário anterior, mas faça alterações rápidas (duas ou três) nos valores monitorados em menos de 3 segundos cada.
2. No OutputTextBox da interface, verifique:
    - O enforcement deve acontecer apenas uma vez em cada janela de 3 segundos.
    - Logs intermediários devem indicar "Enforcement ignorado devido ao debounce" se as alterações forem feitas em sequência muito rápida.

---
## Observações Gerais
- Sempre use a fila "Microsoft Print to PDF" para o MVP, a não ser que parametrize para outras filas.
- Caso apareça algum erro/exceção, o log também será exibido na interface.
- O enforcement só ocorre para os campos ativados na policy (Duplex, Cor, Orientation).
- Se já estiver em conformidade ao detectar o evento, loga "Já em conformidade com a política" e não aplica mudança.

---
Esses cenários garantem validação completa do MVP de enforcement e debounce. Ajuste os valores da policy conforme queira testar outros resultados.
