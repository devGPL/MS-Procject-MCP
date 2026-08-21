# Changelog

## Não publicado

Modularização do servidor e as correções que ela tornou visíveis.

### Mudanças de comportamento

Nenhuma assinatura de ferramenta mudou. O registro MCP é byte a byte idêntico ao da versão em arquivo único — mesma lista de nomes, mesmos schemas de entrada. O que mudou é o que as ferramentas aceitam e o que devolvem em casos que antes falhavam.

**Nomes de calendário passam a casar sem distinguir maiúsculas.**
`set_calendar_exception("standard")` devolvia `{"error": "Calendar 'standard' not found"}` no mesmo projeto em que `set_working_hours("standard")` funcionava. Quatro ferramentas rejeitavam a variação de caixa — `set_calendar_exception`, `set_project_calendar`, `set_task_calendar` e `create_calendar` — enquanto as outras cinco aceitavam. Agora todas casam igual, e a resposta ecoa o nome canônico do calendário.

Junto disso, `create_calendar` passa a recusar um nome que difira apenas na caixa de um já existente. Antes seria possível criar `"standard"` ao lado de `"Standard"`, e `delete_calendar` — que sempre casou sem distinguir caixa — ficaria ambíguo sobre qual dos dois removeu.

**Consultas de tarefa deixam de estourar em campos indisponíveis.**
O helper `safe()` do `task_to_dict` era chamado como `safe(t.ActualStart)`. Python avalia o argumento antes da função rodar, então o `try/except` interno nunca via a falha e o valor padrão era inalcançável. Duas consequências, ambas ativas:

- Uma propriedade COM instável em **uma** tarefa derrubava a listagem inteira. Uma consulta sobre 500 tarefas devolvia erro porque a tarefa 300 tinha `ActualStart` indisponível.
- XML recém-importado devolvia `None` em `RemainingDuration`, `TotalSlack` e `FreeSlack`, e a divisão por `MinutesPerDay` estourava com `TypeError`.

Agora esses campos chegam com o padrão declarado — `remaining_duration_days: 0` em vez de erro, `actual_start: null` em vez de exceção propagada. **Uma tarefa com todos os campos presentes serializa exatamente como antes.**

Continua sem proteção: 18 propriedades são lidas sem `safe` nenhum (`UniqueID`, `Name`, `Duration`, `Start`, `Finish` entre outras). Uma delas falhando ainda derruba a listagem.

**Recurso sem nome no pool deixa de ser fatal.**
`assign_resource`, `get_resource_workload`, `update_resource` e `delete_resource` chamavam `r.Name.lower()` sem verificar. Um único recurso sem nome em qualquer posição do pool fazia as quatro falharem com `AttributeError`, **independentemente de qual recurso estava sendo buscado** — a varredura passa por ele no caminho. Agora recursos sem nome são ignorados.

**Operações em lote passam a usar um snapshot.**
`dry_run_bulk_update` e `bulk_set_deadlines` buscavam cada item varrendo a coleção inteira de tarefas: N itens contra M tarefas custavam N×M travessias COM. Agora montam o mapa uma vez.

A diferença é observável: uma tarefa excluída no meio do mesmo lote continua presente no snapshot. Antes ela sumia da coleção e era reportada como "não encontrada"; agora a referência COM obsoleta falha no uso. Para um lote que exclui e depois toca o mesmo `UniqueID`, o erro muda de "not found" para falha de COM.

### Correções

- O servidor escrevia duas linhas de diagnóstico em **stdout**, que é o transporte stdio do MCP. Agora vão para stderr.
- Três suítes de teste (`test_new_tools`, `test_phase2`, `test_phase3`) não chamavam `sys.exit` e saíam com código 0 mesmo registrando falhas. `test_new_tools` reportava `0 passed, 7 failed` e saía 0.
- Contagens de teste no README estavam erradas em três suítes. São **180 verificações**, não 202.

### Estrutura

- `server.py` deixou de ser um arquivo único de 5.206 linhas e virou um entry point de 55 linhas que importa nove módulos de domínio. Nenhuma ferramenta foi editada durante a divisão.
- `msp_core.py` concentra a fronteira COM. Todo `import win32com` está dentro de corpo de função, então os módulos importam em macOS e Linux sem `pywin32`.
- Lookups reimplementados inline foram consolidados: 13 varreduras de `BaseCalendars` viraram 2, 8 buscas de recurso viraram 1, 8 mapas de `UniqueID` viraram 2.
- Constantes duplicadas (`CONSTRAINT_NAMES`, `TIMESCALE_MAP`) e três cópias locais de formatador de data foram eliminadas.

### Empacotamento

- `pyproject.toml` declara as dependências, que antes só existiam em prosa no README — `python-dateutil` era importado em runtime sem constar em lugar executável.
- Novo comando `msproject-mcp`, o que remove o caminho absoluto da configuração do cliente MCP.
- `mcp` limitado a `<1.20`. A partir da 1.20 ele passou a depender de `pyjwt[crypto]`, que puxa `cryptography` — e `cryptography` publicou wheels para Windows ARM64 só até a 46.0.3. Em Windows on ARM, o `pip install` tentava compilar do zero e falhava por falta do linker do MSVC, exigindo Rust e Visual Studio Build Tools numa arquitetura e nada na outra. Este servidor não autentica nada, então PyJWT era subárvore morta. Com o teto, as 28 dependências têm wheel pronto tanto em `win_arm64` quanto em `win_amd64`, e a instalação é idêntica nas duas.

### Ferramentas de desenvolvimento

`tools/gates.sh` roda três portões em menos de meio segundo, sem Windows e sem MS Project:

| Portão | Pega |
|---|---|
| `check_names.py` | nome carregado sem nada que o defina |
| `snap_tools.py` | ferramenta que sumiu do registro, comparando por nome |
| `check_unused.py` | import morto, banner de seção sem código |

O segundo existe porque as ferramentas se registram por efeito colateral de import: remover um `import` de módulo apaga as ferramentas dele do registro sem `ImportError` nenhum.
