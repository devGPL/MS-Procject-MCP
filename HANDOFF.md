# Handoff — sessão de 21/08/2026

Estado do projeto, medições que não precisam ser refeitas, e o que ficou aberto.

---

## 1. Estado

### Mergeado em `main`

| PR | O quê |
|---|---|
| #2 | README em pt-BR, guias de configuração Claude e Codex |
| #3 | Modularização: `server.py` de 5.206 linhas → 10 módulos + entry point de 55 |
| #4 | Aviso quando não há rede de dependências para CPM |
| #5 | Retry de `RPC_E_CALL_REJECTED` (MS Project ocupado) |

### Aberto

**PR #6** — detecta agendamento manual, não só ausência de ligações. Verificado contra o cronograma real. Pronto para revisão.

### Não enviado

Branch `claude/reduce-com-traversals`, três commits:

- `2fc5704` — corta travessias COM redundantes (`get_project_info` 4→1; supressão de recálculo em 4 tools de lote)
- `faf9909` — todas as tools mutantes passam a salvar
- `98dbc9a` — caminho rápido via `mpxj`, integrado em `get_critical_path`

O terceiro depende dos dois anteriores: sem o `faf9909`, uma escrita não salva ficaria invisível para a leitura seguinte.

### Ambiente

- Servidor roda numa VM Windows ARM64 (Parallels, `10.211.55.5`), HTTP na porta 8765
- Claude Code e Claude Desktop configurados no Mac; Desktop via ponte `mcp-remote`
- Backup do cronograma: `~/Downloads/Cronograma_R4A_Adaptado_Thomas.BACKUP-20260821-172102.mpp`

---

## 2. As três barreiras de COM

Todas dão o **mesmo** erro ou nenhum. Vale checar nesta ordem antes de investigar qualquer outra coisa.

| # | Condição | Sintoma |
|---|---|---|
| 1 | Sessão de logon | SSH cai na sessão 0, MS Project vive na 1. `MK_E_UNAVAILABLE` |
| 2 | Nível de integridade | Servidor elevado (admin) não enxerga MS Project normal. Mesmo erro |
| 3 | Aplicativo ocupado | `RPC_E_CALL_REJECTED`. Transitório, resolve sozinho |

A **2** é a mais cruel: o instinto diante de um erro de acesso é rodar como administrador, e é exatamente isso que quebra. Documentado em `msp_core.get_app()`, no README e na seção de VM.

Consequência prática: o servidor **precisa** subir da área de trabalho da VM, sem elevação. Nunca por SSH.

---

## 3. Medições — não refazer

Cronograma de referência: `Cronograma_R4A_Adaptado_Thomas.mpp`, 8.429 tarefas, 13,8 MB.

### Custo de COM

| | |
|---|---|
| Enumerar uma tarefa | **2,635 ms** |
| Ler uma propriedade | **0,240 ms** |
| Enumeração como fatia do custo | **83%** |
| Piso para varrer 8.429 tarefas | **~20s** |

**Otimizar leitura de propriedade ataca 17% do problema.** O caro é obter o objeto.

Descartados por medição:
- **Paralelizar**: COM é STA. 2 chamadas simultâneas = 37,7s vs 19,2s de uma. Serializa 1:1
- **Acesso indexado `Tasks(i)`**: 97,9s projetado vs 26,5s do enumerador. 4x pior
- **Early binding / `gencache`**: as duas APIs empatam. O ganho aparente de 4,37x era cache de tipos ausente vs presente, não escolha de API

### Tempos das ferramentas (COM, cronograma real)

`get_critical_path` 29,8s · `get_schedule_analysis` 43,1s · `find_available_slack` 40,4s · `what_if_delay` 48,0s

Variância entre execuções idênticas chega a **4,9s** — diferenças abaixo disso não são mensuráveis com 2 rodadas.

### Caminho rápido (`mpxj`)

| | |
|---|---|
| Parse de 13,8 MB | 1,42s |
| 8.243 tarefas, dict de 35 campos | 0,55s |
| Segunda leitura (cache por mtime) | 0,000s |

Paridade com COM confirmada: 8.243 de detalhe, 61 críticas, 7.911 ligadas, 0 manuais, mesmas 35 chaves.

**Indisponível em Windows ARM64**: `jpype1` não publica wheel. Requer Java 9+.

### Payload — medido, não resolvido

| | |
|---|---|
| `get_tasks` completo, JSON `indent=2` | **8,07 MB** |
| JSON compacto | 5,79 MB (72%) |
| CSV | 1,99 MB (25%) |
| Valores vazios/zero | **56%** |
| Definições das 99 tools | 51 KB (~13k tokens) |

Os dois últimos consomem janela de contexto. **Não aparecem como lentidão — aparecem como conversa que acaba cedo.**

---

## 4. O cronograma do cliente

Dois arquivos analisados, ambos vindos de planejamento que não usa CPM:

| | `V07 TESTE` | `Thomas` (antes da conversão) |
|---|---|---|
| Tarefas de detalhe | 7.583 | 8.243 |
| Com predecessora | **0** | 7.911 (96%) |
| Agendamento manual | 100% | 100% |
| Críticas de detalhe | 2 | 0 |

**Conversão executada em 21/08 no `Thomas`**: todas as de detalhe passaram para automático. O fim do projeto saltou de **2028-06-30 para 2031-03-05** — 2 anos e 8 meses. As datas manuais afirmavam um cronograma que a própria rede de dependências não sustenta.

Isso é informação de negócio. O backup tem o estado anterior.

---

## 5. Portões

`./tools/gates.sh` — 0,43s, sem Windows e sem MS Project:

| Portão | Pega |
|---|---|
| `check_names.py` | nome carregado sem nada que o defina |
| `snap_tools.py` | ferramenta sumida do registro, por nome |
| `check_unused.py` | import morto, banner órfão |

O sha256 do registro (`5af9975e…`) foi congelado antes do primeiro corte e **não mudou em nenhum commit da modularização**.

`tests/test_backend_parity.py` compara COM × `mpxj` campo a campo. **Nunca rodou** — exige Windows x64 com projeto salvo, e a VM é ARM64.

---

## 6. Aberto

**Alto valor**

- Replicar o caminho rápido nas outras 10 ferramentas de leitura pesada (padrão já estabelecido em `get_critical_path`)
- Rodar `test_backend_parity.py` numa máquina x64 — a paridade nunca foi provada em teste, só observada
- Payload: `indent=2` → compacto é barato e mede-se na hora

**Médio**

- As 180 verificações de integração nunca rodaram nesta série
- 18 propriedades em `task_to_dict` são lidas sem `safe()`
- 15 das 99 tools não são exercitadas por nenhuma suíte, incluindo `add_task` e `delete_task`
- 4 tools confirmam operações que podem não ter acontecido (`undo_last`, `reschedule_incomplete_work`, `set_working_hours`, `get_task_dependencies`)

**Descartado, com motivo**

- **CData MCP** — exige Project Online, só leitura, driver licenciado, Java. O desenho de 3 tools genéricas só funciona porque há motor SQL do outro lado
- **`mpxj` na VM atual** — ARM64 não tem `jpype1`

---

## 7. Método — a lição que mais custou

**Errei quatro vezes seguidas pelo mesmo motivo: medi o pedaço que eu já suspeitava, em vez do caminho inteiro.**

| Hipótese | Como caiu |
|---|---|
| "É ligação tardia, 4,37x" | Segundo benchmark: as duas APIs empatam |
| "É o cache de tipos, reinicia" | Reiniciou, continuou em 25s |
| "Os dois cronogramas não têm malha" | Extrapolei de um para o outro. O usuário corrigiu |
| "A sonda detecta o problema" | Ela checava predecessora; a causa era agendamento manual |

Em todos, o benchmark que acertou foi o primeiro que cronometrou **exatamente o que o servidor faz**, sem eu escolher antes o que contava. O primeiro deles excluía a enumeração do relógio de propósito — "para isolar a variável" — e isolou a errada.

Outros erros da mesma família:

- `pgrep -x Claude` não achou o app (nome exato não bate) → concluí "fechado" e editei a config por baixo dele
- `MainWindowTitle` consultado por SSH da sessão 0 → vem vazio sempre; li como "sem projeto aberto"
- Estimei linhas em mensagem de commit em vez de medir. Três vezes
- Usei `net.sf.mpxj` de memória; o pacote virou `org.mpxj` na versão 14

**Regra para a próxima:** quando uma medição contradiz o modelo, o modelo está errado até prova em contrário — e a prova é medir o caminho completo, não o pedaço suspeito. E quando o usuário contradiz uma afirmação minha, medir antes de argumentar; ele estava certo nas duas vezes.

**O que funcionou:** testar cada portão contra quebra sintética antes de confiar nele. O Gate 3 deu falso positivo na primeira execução (marcou `# COM helpers` do `msp_core` como banner órfão) e isso só apareceu porque rodei antes de commitar.
