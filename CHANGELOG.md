# Changelog

## Não publicado

Modularização do servidor e as correções que ela tornou visíveis.

### Página do PyPI

**O PyPI passa a mostrar um README próprio, só de instalação.** A página do pacote exibia o README completo do repositório (456 linhas, incluindo detalhes do ambiente de desenvolvimento). O `pyproject.toml` agora aponta para `README-pypi.md`: requisitos, `pip install msproject-mcp`, extra `[fast]`, registro nos clientes MCP e link para a documentação completa no GitHub. O README do repositório segue intacto como doc autoritativa. Versão vai a 0.7.1 porque a descrição no PyPI é congelada por versão publicada.

### Tamanho das respostas

**`get_tasks` devolvia 8,07 MB num cronograma real; passa a devolver ~82 KB.** Medido ponta a ponta com 8.243 tarefas: `filter_tasks` 7,66 MB → 81,7 KB, `get_schedule_analysis` 1,93 MB → 37,4 KB, `get_wbs_structure` 3,51 MB → 43,6 KB.

Três mudanças de comportamento, todas visíveis para quem consome as respostas:

- **Listagens param em 200 itens.** `count` e `total` continuam completos; um bloco `page` traz `returned`, `offset` e como pedir o resto. Sem teto: `filter_tasks` com `limit=-1` (o default de `limit` passou de "tudo" para 200) e `export_csv`, que escreve arquivo.
- **Campos vazios não são publicados.** Chave ausente significa vazio, zero ou falso. Preservados: `unique_id`, `id`, `name`, e os números onde zero é medição (`percent_complete`, `total_slack_days`, `free_slack_days`, `duration_days`, `outline_level`).
- **`get_wbs_structure` passa a cortar em `max_level=3`** por padrão, dizendo quantas tarefas ficaram de fora. `0` devolve a árvore inteira.

Respostas acima de 4 KB saem em JSON compacto; abaixo disso seguem indentadas.

**As definições das 99 ferramentas caíram de 50,7 KB para 37,5 KB** (~14,4k → ~10,7k tokens), custo que todo cliente paga em toda sessão antes da primeira chamada: o `"title"` que o pydantic gera para cada parâmetro saiu do schema publicado (8 KB), e as descrições das 40 maiores foram reescritas mantendo enums, formatos, unidades e sentinelas.

### Leitura

**Dezesseis ferramentas de leitura passam a responder do arquivo salvo quando podem.** Antes só `get_critical_path` fazia isso. Agora também `get_tasks`, `get_task`, `search_tasks`, `get_tasks_by_rag`, `get_overdue_tasks`, `get_tasks_by_resource`, `get_progress_summary`, `get_wbs_structure`, `filter_tasks`, `group_tasks_by`, `get_progress_by_wbs`, `get_constraints`, `export_csv`, `get_schedule_analysis`, `find_available_slack` e `validate_schedule`. Toda resposta ganha um bloco `source` dizendo qual caminho respondeu e por quê — e um `warning` quando o MS Project reporta alterações não salvas. Requer `pip install -e ".[fast]"`; sem isso, ou em Windows ARM64, tudo cai no COM e o `source` explica.

Enumerar uma tarefa pelo COM custa 2,635 ms contra 0,240 ms para ler uma propriedade dela: 83% de uma varredura acontece antes de qualquer campo ser lido. Numa agenda de 8.429 tarefas isso é um piso de ~20 s por ferramenta.

Duas mudanças estreitas de comportamento acompanham:

- `group_tasks_by` com um nome de campo fora da lista documentada agora o procura entre os campos publicados de uma tarefa (as chaves que `get_task` devolve) em vez de entre as propriedades COM cruas.
- `get_constraints` reporta um valor de restrição desconhecido como ausente, em vez de `"Unknown(9)"`. Os oito valores conhecidos são o enum inteiro.

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

**A proteção passou a cobrir os campos onde um default significa "ausente".**
Antes metade da tabela de leitores era defensiva e metade não, sem razão declarada. Agora `Start`, `Finish`, `WBS`, `Predecessors`, `ResourceNames`, `Notes`, `Deadline`, `Priority`, as folgas, os textos, os flags, `Milestone`, `Critical` e `Active` sobrevivem a uma leitura COM que falhe, cada um com seu padrão declarado — uma tarefa ilegível custa aquele campo, não a varredura inteira.

Sete leitores continuam **crus de propósito**, e a tabela agora diz por quê: identidade (`unique_id`, `id`, `name`), estrutura (`outline_level`, `summary`) e medição (`percent_complete`, `duration_days`). Nesses, um default não seria "ausente" e sim uma afirmação sobre o projeto — identidade inventada publica a tarefa errada, estrutura inventada remonta a árvore WBS, e `0` de medição vira achado do `validate_schedule`. Uma falha ali deve chegar ao chamador, não virar número plausível.

No caminho de sucesso, quatro campos guardados se aproximam do backend de arquivo: `Predecessors`, `Notes`, `ResourceNames` e `WBS` publicam `""` onde o COM devolvia `None`, e `active` cai para `True` em vez de `False` — o mpxj já respondia assim, então os dois backends concordam em mais coisa.

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
- A detecção de órfãs do `validate_schedule` deixou de ser O(n²) em leituras COM. Para cada tarefa de detalhe ela relia `Predecessors` de todas as outras: ~71 milhões de leituras numa agenda de 8.429 tarefas, o bastante para tornar a ferramenta inutilizável ali. Uma passada monta o mapa `linha → quem a cita` e o teste vira consulta de pertinência. No caminho, um bug de locale: o separador da lista de predecessores é `,` numa instalação inglesa e `;` numa portuguesa (e o backend de arquivo usa `;` sempre), então dividir só por um deles não achava sucessor nenhum no outro — lido como "toda tarefa é órfã". O parser agora aceita os dois.
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
