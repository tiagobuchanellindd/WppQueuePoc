# WPP Queue API / POC Documentation
---
## EN
This repository documents a Windows printing API/POC focused on **Windows Protected Print (WPP)** behavior, queue lifecycle management, and native Winspool interoperability.
### What this project covers
- Global WPP status detection from Windows Registry.
- Queue lifecycle operations: create, list, update, delete.
- Queue inspection heuristics (`LikelyWpp`, `LikelyNotWpp`, `Indeterminate`).
- Print ticket read/update flows (default and user scope).
- Native interop patterns (`winspool.drv`, P/Invoke, buffer management, error handling).
### Documentation structure
- `docs/guides/`
  - `docs/guides/01-objectives-summary.md`
  - `docs/guides/02-business-overview.md`
  - `docs/guides/03-architecture-styles.md`
  - `docs/guides/04-native-methods-guide.md`
- `docs/estudo/`
  - In-depth study notes (PT-BR), organized by source file.
### Recommended reading order
1. `docs/guides/01-objectives-summary.md` — delivery goals and scope  
2. `docs/guides/02-business-overview.md` — business context and expected behavior  
3. `docs/guides/03-architecture-styles.md` — architecture and layer interaction  
4. `docs/guides/04-native-methods-guide.md` — Winspool interop details  
5. `docs/estudo/README.md` — PT-BR study index and deep method-level notes  
---
## PT-BR
Este repositório documenta uma API/POC de impressão no Windows, com foco no comportamento do **Windows Protected Print (WPP)**, no ciclo de vida de filas e na interoperabilidade nativa com Winspool.
### O que este projeto cobre
- Detecção do status global de WPP via Registro do Windows.
- Operações de ciclo de vida de fila: criar, listar, atualizar e excluir.
- Heurística de inspeção de fila (`LikelyWpp`, `LikelyNotWpp`, `Indeterminate`).
- Leitura/atualização de print ticket (escopo default e user).
- Padrões de interop nativo (`winspool.drv`, P/Invoke, buffers, tratamento de erros).
### Estrutura da documentação
- `docs/guides/`
  - `docs/guides/01-objectives-summary.md`
  - `docs/guides/02-business-overview.md`
  - `docs/guides/03-architecture-styles.md`
  - `docs/guides/04-native-methods-guide.md`
- `docs/estudo/`
  - Estudos aprofundados (PT-BR), organizados por arquivo-fonte.
### Ordem recomendada de leitura
1. `docs/guides/01-objectives-summary.md` — objetivos de entrega e escopo  
2. `docs/guides/02-business-overview.md` — contexto de negócio e comportamento esperado  
3. `docs/guides/03-architecture-styles.md` — arquitetura e interação entre camadas  
4. `docs/guides/04-native-methods-guide.md` — detalhes de interop com Winspool  
5. `docs/estudo/README.md` — índice de estudos e trilhas detalhadas por método  
