"""
The two read backends must answer identically. This proves it offline.

Twelve reading tools now take whatever msp_fast.varredura hands them: dicts
parsed from the saved .mpp, or COM tasks wrapped in a view that answers to the
same key names. One body, two sources -- and a body that reads correctly from
one and subtly wrong from the other is exactly the failure the fast path makes
possible, because the wrong answer is plausible and nothing marks it.

So every tool here runs TWICE over the same fake project: once forced down the
COM branch, once with the parsed list handed to it, and the two responses must
match field for field. It also counts COM property reads, which is the second
claim the design rests on -- the view reads a property only when a tool asks
for it, so a tool that publishes three fields must not pay for thirty-five.

This suite needs NO Windows and NO Microsoft Project: the COM side is a fake.
That is its limit as well as its point. It proves the two BODIES agree; it
cannot prove that mpxj and Microsoft Project read the same file the same way
(tests/test_backend_parity.py does that, on Windows). Note also that a Python
fake raises AttributeError where real COM raises com_error, so defensive reads
that pass here can still fail in production.

    python tests/test_vistas_paridade.py
"""

import asyncio
import datetime
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import msp_core
import msp_fast
import msp_schedule
import msp_tasks_read
from server import mcp

PASS = 0
FAIL = 0

MPD = 480

# Rounded down to the hour so that two runs minutes apart build the SAME
# fixture: comparing this suite's output against another build's is worthless
# if every date differs by a second.
AGORA = datetime.datetime.now().replace(minute=0, second=0, microsecond=0)
ONTEM = AGORA - datetime.timedelta(days=30)
AMANHA = AGORA + datetime.timedelta(days=30)

# Every property task_to_dict reads, with a default -- so a fixture declares
# only what the case is about.
PADROES = {
    "UniqueID": 0, "ID": 0, "Name": "", "OutlineLevel": 1, "WBS": "",
    "Summary": False, "Milestone": False, "Start": None, "Finish": None,
    "Duration": 480, "PercentComplete": 0, "ActualStart": None,
    "ActualFinish": None, "RemainingDuration": 0, "TotalSlack": 0,
    "FreeSlack": 0, "Deadline": None, "Priority": 500, "ConstraintType": 0,
    "ConstraintDate": None, "Manual": False, "Type": 0, "Predecessors": "",
    "ResourceNames": "", "Notes": "", "Critical": False, "Active": True,
    "Text1": "", "Text2": "", "Text3": "", "Flag1": False, "Flag2": False,
    "HyperlinkAddress": "", "Hyperlink": "",
}


class TarefaFalsa:
    """A COM task with a counter on every property read."""

    def __init__(self, contador, **campos):
        object.__setattr__(self, "_valores", dict(PADROES, **campos))
        object.__setattr__(self, "_contador", contador)

    def __getattr__(self, nome):
        valores = object.__getattribute__(self, "_valores")
        if nome not in valores:
            raise AttributeError(nome)
        object.__getattribute__(self, "_contador")[0] += 1
        return valores[nome]


class TarefaQueFalha(TarefaFalsa):
    """Uma tarefa COM em que ler certas propriedades levanta.

    O COM real levanta com_error onde este fake levanta RuntimeError. O tipo
    nao importa aqui: _safe captura Exception, e o que esta sob teste e QUAIS
    campos ele protege, nao de que classe e a falha.
    """

    def __init__(self, contador, falham=(), **campos):
        super().__init__(contador, **campos)
        object.__setattr__(self, "_falham", set(falham))

    def __getattr__(self, nome):
        if nome in object.__getattribute__(self, "_falham"):
            raise RuntimeError("com_error simulado ao ler %s" % nome)
        return super().__getattr__(nome)


class ProjetoFalso:
    def __init__(self, tarefas, nome="Projeto Falso"):
        self.Tasks = tarefas
        self.Name = nome
        self.MinutesPerDay = MPD
        self.ProjectStart = datetime.datetime(2026, 1, 1)
        self.ProjectFinish = datetime.datetime(2027, 1, 1)


class AplicacaoFalsa:
    def __init__(self, proj):
        self.ActiveProject = proj


def projeto_de_teste(contador):
    """A project with the shapes the tools branch on.

    Summaries and detail, a milestone, an overdue task, one manual task, one
    linked, a constraint, RAG values, slack above and below the default
    threshold -- so that no tool under test filters everything away and
    reports parity on two empty lists.
    """
    def t(**campos):
        return TarefaFalsa(contador, **campos)

    return [
        t(UniqueID=1, ID=1, Name="Programa", OutlineLevel=1, Summary=True,
          WBS="1", Start=datetime.datetime(2026, 1, 5), Finish=AMANHA,
          Duration=9600, PercentComplete=40),
        t(UniqueID=2, ID=2, Name="Requisitos", OutlineLevel=2, WBS="1.1",
          Start=datetime.datetime(2026, 1, 5), Finish=ONTEM,
          Duration=4800, PercentComplete=0, Text1="Red", Critical=True,
          ResourceNames="Alice", TotalSlack=0, FreeSlack=0,
          Predecessors="", Manual=False),
        t(UniqueID=3, ID=3, Name="Desenho", OutlineLevel=2, WBS="1.2",
          Start=datetime.datetime(2026, 2, 2), Finish=AMANHA,
          Duration=7200, PercentComplete=50, Text1="Amber",
          ResourceNames="Bob, Alice", TotalSlack=4800, FreeSlack=2400,
          Predecessors="2FS", Manual=False, ConstraintType=4,
          ConstraintDate=datetime.datetime(2026, 2, 2), Flag1=True),
        t(UniqueID=4, ID=4, Name="Marco de Revisao", OutlineLevel=2, WBS="1.3",
          Milestone=True, Duration=0, Start=ONTEM, Finish=ONTEM,
          PercentComplete=0, Text1="Green", Predecessors="3FS", Manual=True),
        t(UniqueID=5, ID=5, Name="Entrega", OutlineLevel=1, Summary=True,
          WBS="2", Start=datetime.datetime(2026, 3, 1), Finish=AMANHA,
          Duration=9600, PercentComplete=10),
        t(UniqueID=6, ID=6, Name="Construcao", OutlineLevel=2, WBS="2.1",
          Start=datetime.datetime(2026, 3, 1), Finish=AMANHA,
          Duration=14400, PercentComplete=100, Text1="", Critical=True,
          TotalSlack=0, Predecessors="4FS", Manual=False,
          ResourceNames="Carol", Text2="Fase 2", Notes="nota"),
        t(UniqueID=7, ID=7, Name="Treinamento", OutlineLevel=2, WBS="2.2",
          Start=datetime.datetime(2026, 4, 1), Finish=AMANHA,
          Duration=2400, PercentComplete=0, TotalSlack=9600, FreeSlack=4800,
          Manual=True, Active=False),
        # Two predecessors in ONE string, and the second of them is task 7 --
        # which has no predecessors of its own. So 7 is a successor only if the
        # orphan scan splits the entry: a scan that reads "6FS,7FS" whole, or
        # keeps only the first entry, calls 7 an orphan and is wrong. Missing
        # start and negative slack here give validate_schedule two more
        # categories to fill, so it does not answer with empty lists that any
        # two backends agree on.
        t(UniqueID=8, ID=8, Name="Comissionamento", OutlineLevel=2, WBS="2.3",
          Start=None, Finish=AMANHA, Duration=2400, PercentComplete=0,
          TotalSlack=-2400, FreeSlack=0, Predecessors="6FS,7FS",
          ResourceNames="Dave", Manual=False),
        # Nobody links to it and it links to nobody: the real orphan.
        t(UniqueID=9, ID=9, Name="Auditoria", OutlineLevel=2, WBS="2.4",
          Start=datetime.datetime(2026, 5, 1), Finish=AMANHA, Duration=2400,
          PercentComplete=0, ResourceNames="Erin", Manual=False),
        # A summary with nothing under it: last row, so no task follows that
        # could be its child.
        t(UniqueID=10, ID=10, Name="Encerramento", OutlineLevel=1, Summary=True,
          WBS="3", Start=datetime.datetime(2026, 6, 1), Finish=AMANHA,
          Duration=480, PercentComplete=0),
    ]


def sem_origem(valor):
    """The response with every 'source' block stripped, at any depth."""
    if isinstance(valor, dict):
        return {k: sem_origem(v) for k, v in valor.items() if k != "source"}
    if isinstance(valor, list):
        return [sem_origem(v) for v in valor]
    return valor


# --- o que a tabela de leitores deve produzir ----------------------------
#
# The parity checks above compare the two BODIES, and both get their fields
# from the same table of readers -- so a reader that returns the wrong thing
# is wrong on both sides and they still agree. These expectations are written
# out by hand, against the fixture rather than against the code, and are the
# only check here that a field means what it says.

CAMPOS_ESPERADOS = (
    "unique_id", "id", "name", "outline_level", "wbs", "summary", "milestone",
    "start", "finish", "duration_days", "percent_complete", "actual_start",
    "actual_finish", "remaining_duration_days", "total_slack_days",
    "free_slack_days", "deadline", "priority", "constraint_type",
    "constraint_date", "manual", "type", "predecessors", "resource_names",
    "notes", "critical", "active", "rag", "text1", "text2", "text3",
    "flag1", "flag2", "hyperlink", "hyperlink_text",
)

# Keyed by UniqueID; only the fields the converted tools branch on.
ESPERADO = {
    2: {"name": "Requisitos", "summary": False, "milestone": False,
        "critical": True, "manual": False, "active": True, "text1": "Red",
        "rag": "Red", "percent_complete": 0, "outline_level": 2,
        "total_slack_days": 0.0, "free_slack_days": 0.0, "duration_days": 10.0,
        "constraint_type": "ASAP", "predecessors": "", "resource_names": "Alice"},
    3: {"name": "Desenho", "summary": False, "critical": False,
        "manual": False, "text1": "Amber", "percent_complete": 50,
        "total_slack_days": 10.0, "free_slack_days": 5.0, "flag1": True,
        "constraint_type": "SNET", "constraint_date": "2026-02-02 00:00:00",
        "predecessors": "2FS", "resource_names": "Bob, Alice",
        "duration_days": 15.0},
    4: {"name": "Marco de Revisao", "milestone": True, "manual": True,
        "duration_days": 0, "text1": "Green", "critical": False,
        "predecessors": "3FS"},
    6: {"name": "Construcao", "critical": True, "percent_complete": 100,
        "manual": False, "text2": "Fase 2", "notes": "nota",
        "resource_names": "Carol", "total_slack_days": 0.0,
        "duration_days": 30.0, "type": "FixedUnits", "priority": 500},
    7: {"name": "Treinamento", "manual": True, "active": False,
        "total_slack_days": 20.0, "free_slack_days": 10.0,
        "percent_complete": 0, "summary": False},
    8: {"name": "Comissionamento", "summary": False, "start": None,
        "predecessors": "6FS,7FS", "resource_names": "Dave",
        "total_slack_days": -5.0, "percent_complete": 0},
    9: {"name": "Auditoria", "summary": False, "predecessors": "",
        "resource_names": "Erin", "total_slack_days": 0.0},
    10: {"name": "Encerramento", "summary": True, "outline_level": 1,
         "wbs": "3", "duration_days": 1.0},
}


async def chamar(nome, **kwargs):
    resultado = await mcp.call_tool(nome, kwargs)
    conteudo = resultado[0] if isinstance(resultado, tuple) else resultado
    texto = conteudo[0].text if conteudo else ""
    return json.loads(texto) if texto else {}


def instalar_falsos(proj):
    """Point every converted module at the fake project."""
    for mod in (msp_tasks_read, msp_schedule):
        mod.get_app = lambda: AplicacaoFalsa(proj)
        mod.get_proj = lambda app: app.ActiveProject


async def rodar():
    global PASS, FAIL

    def ok(nome, cond, detalhe=""):
        global PASS, FAIL
        if cond:
            PASS += 1
            print("  PASS  %s" % nome)
        else:
            FAIL += 1
            print("  FAIL  %s  %s" % (nome, detalhe))

    contador = [0]
    tarefas = projeto_de_teste(contador)
    proj = ProjetoFalso(tarefas)
    instalar_falsos(proj)

    # What a perfect parse of this project would hand back: the same fields,
    # in the row order the fast path sorts into.
    lidas = sorted((msp_core.task_to_dict(t, MPD) for t in tarefas),
                   key=lambda d: d["id"])

    csv_com  = os.path.join(tempfile.gettempdir(), "paridade_com.csv")
    csv_arq  = os.path.join(tempfile.gettempdir(), "paridade_arquivo.csv")

    CASOS = [
        ("get_tasks", {}),
        ("get_tasks", {"include_summary": True}),
        ("get_tasks", {"outline_level": 2}),
        ("get_tasks", {"keyword": "de"}),
        ("search_tasks", {"query": "marco"}),
        ("get_task", {"unique_id": 6}),
        ("get_task", {"unique_id": 999}),
        ("get_tasks_by_rag", {"rag": "Red"}),
        ("get_tasks_by_rag", {"rag": "amber"}),
        ("get_overdue_tasks", {}),
        ("get_tasks_by_resource", {"resource_name": "alice"}),
        ("get_progress_summary", {}),
        ("get_wbs_structure", {}),
        ("get_wbs_structure", {"max_level": 2}),
        ("filter_tasks", {"filters_json": json.dumps({"rag": "Red"})}),
        ("filter_tasks", {"filters_json": json.dumps({"critical": True, "summary": False})}),
        ("filter_tasks", {"filters_json": json.dumps({"start_after": "2026-02-01"})}),
        ("filter_tasks", {"filters_json": json.dumps({"finish_before": "2030-01-01", "min_pct": 50})}),
        ("filter_tasks", {"filters_json": json.dumps({"resource": "bob", "sort_by": "name"})}),
        ("filter_tasks", {"filters_json": json.dumps({"active": False})}),
        ("filter_tasks", {"filters_json": json.dumps({"name_contains": "entrega", "limit": 1})}),
        ("group_tasks_by", {"field": "rag"}),
        ("group_tasks_by", {"field": "resource", "include_tasks": True}),
        ("group_tasks_by", {"field": "percent_complete"}),
        ("group_tasks_by", {"field": "outline_level"}),
        ("group_tasks_by", {"field": "nao_existe"}),
        ("get_progress_by_wbs", {}),
        ("get_progress_by_wbs", {"max_level": 1}),
        ("get_constraints", {}),
        ("get_critical_path", {}),
        ("get_schedule_analysis", {}),
        ("find_available_slack", {}),
        ("find_available_slack", {"min_days": 20}),
        ("validate_schedule", {}),
    ]

    print("\n=== PARIDADE ENTRE OS DOIS CAMINHOS ===")

    for nome, kwargs in CASOS:
        msp_fast.varredura = lambda p: (None, {"backend": "com", "reason": "test"})
        antes = contador[0]
        via_com = await chamar(nome, **kwargs)
        leituras_com = contador[0] - antes

        msp_fast.varredura = lambda p, _l=lidas: ([dict(d) for d in _l],
                                                  {"backend": "mpxj"})
        antes = contador[0]
        via_arquivo = await chamar(nome, **kwargs)
        leituras_arquivo = contador[0] - antes

        rotulo = "%s %s" % (nome, json.dumps(kwargs, sort_keys=True))
        ok(rotulo, sem_origem(via_com) == sem_origem(via_arquivo),
           "COM=%s\n        ARQUIVO=%s"
           % (json.dumps(sem_origem(via_com))[:400],
              json.dumps(sem_origem(via_arquivo))[:400]))
        ok(rotulo + " -- caminho rapido nao toca COM", leituras_arquivo == 0,
           "leu %d propriedades COM" % leituras_arquivo)
        if leituras_com == 0 and nome != "get_task":
            ok(rotulo + " -- caminho COM leu algo", False,
               "nenhuma propriedade lida; o teste nao exercitou nada")

    # export_csv writes a file, so it is compared by content rather than by
    # the response, which carries a different path in each run.
    msp_fast.varredura = lambda p: (None, {"backend": "com", "reason": "test"})
    await chamar("export_csv", output_path=csv_com)
    msp_fast.varredura = lambda p, _l=lidas: ([dict(d) for d in _l], {"backend": "mpxj"})
    await chamar("export_csv", output_path=csv_arq)
    ok("export_csv -- mesmo CSV pelos dois caminhos",
       open(csv_com, encoding="utf-8").read() == open(csv_arq, encoding="utf-8").read())

    # The claim the whole view design rests on: a tool that publishes a few
    # fields must not pay for all thirty-five. get_progress_summary reads five.
    print("\n=== CUSTO DA VISTA COM ===")
    msp_fast.varredura = lambda p: (None, {"backend": "com", "reason": "test"})
    antes = contador[0]
    await chamar("get_progress_summary", **{})
    por_tarefa = (contador[0] - antes) / float(len(tarefas))
    ok("get_progress_summary le menos de 10 propriedades por tarefa",
       por_tarefa < 10, "leu %.1f" % por_tarefa)

    antes = contador[0]
    await chamar("get_tasks", include_summary=True)
    por_tarefa = (contador[0] - antes) / float(len(tarefas))
    ok("get_tasks publica tudo e le as 35", 34 <= por_tarefa <= 36,
       "leu %.1f" % por_tarefa)

    # A view must answer to exactly the keys task_to_dict publishes, or a body
    # written against one backend reads None from the other and says nothing.
    print("\n=== CONTRATO DA VISTA ===")
    v = msp_core.VistaCOM(tarefas[2], MPD)
    d = msp_core.task_to_dict(tarefas[2], MPD)
    ok("VistaCOM cobre as mesmas chaves",
       set(msp_core.CAMPOS_TAREFA) == set(d))
    ok("VistaCOM devolve os mesmos valores",
       all(v[k] == d[k] for k in d),
       str([k for k in d if v[k] != d[k]]))
    ok("VistaCOM.completo() == task_to_dict", v.completo() == d)
    ok("tarefa_completa e identidade para dict",
       msp_core.tarefa_completa(d) is d)
    try:
        v["campo_inexistente"]
        ok("chave inexistente levanta KeyError", False)
    except KeyError:
        ok("chave inexistente levanta KeyError", True)

    print("\n=== O QUE OS LEITORES DEVEM DEVOLVER ===")
    ok("a tabela publica exatamente os 35 campos esperados",
       list(msp_core.CAMPOS_TAREFA) == list(CAMPOS_ESPERADOS),
       "sobrando %s / faltando %s"
       % (sorted(set(msp_core.CAMPOS_TAREFA) - set(CAMPOS_ESPERADOS)),
          sorted(set(CAMPOS_ESPERADOS) - set(msp_core.CAMPOS_TAREFA))))

    por_uid = {msp_core.task_to_dict(t, MPD)["unique_id"]:
               msp_core.task_to_dict(t, MPD) for t in tarefas}
    for uid, esperado in ESPERADO.items():
        lido = por_uid[uid]
        erradas = {k: (v, lido.get(k)) for k, v in esperado.items()
                   if lido.get(k) != v}
        ok("tarefa %d lida campo a campo" % uid, not erradas, str(erradas))


    # --- o que validate_schedule deve encontrar --------------------------
    #
    # Parity runs the SAME body twice, so an orphan scan that is wrong is
    # wrong on both sides and the two still agree. These expectations are
    # written by hand against the fixture, and are the only check that the
    # categories mean what they say.
    print("\n=== O QUE validate_schedule DEVE ENCONTRAR ===")
    msp_fast.varredura = lambda p: (None, {"backend": "com", "reason": "test"})
    saida = await chamar("validate_schedule")

    def nomes(categoria):
        return sorted(x["name"] for x in saida["issues"][categoria]["tasks"])

    ESPERADAS = {
        # 7 has no predecessors and is named only as the SECOND entry of
        # "6FS,7FS", so it is an orphan to any scan that fails to split.
        "orphan_tasks":         ["Auditoria"],
        "no_resources":         ["Treinamento"],
        "past_due_no_progress": ["Marco de Revisao", "Requisitos"],
        "empty_summaries":      ["Encerramento"],
        "missing_dates":        ["Comissionamento"],
        "negative_slack":       ["Comissionamento"],
    }
    for categoria, esperadas in ESPERADAS.items():
        ok("validate_schedule %s" % categoria, nomes(categoria) == esperadas,
           "achou %s" % nomes(categoria))
    ok("validate_schedule conta e pontua",
       saida["summary"] == {"total_tasks": 10, "total_issues": 7}
       and saida["health_score"] == 30,
       json.dumps({"summary": saida["summary"],
                   "health_score": saida["health_score"]}))

    # --- e a que custo ---------------------------------------------------
    #
    # The orphan scan used to re-read Predecessors of every task for every
    # detail task: O(n^2) COM reads, about 71 million on the reference
    # 8,429-task schedule. A ten-task fixture cannot tell n^2 from n, so this
    # runs on a project large enough that it can: at 120 tasks the old shape
    # costs upwards of 120 reads per task and the one-pass shape costs a
    # dozen.
    print("\n=== CUSTO DE validate_schedule ===")
    contador_g = [0]
    N = 120
    grandes = [TarefaFalsa(contador_g, UniqueID=i, ID=i, Name="T%d" % i,
                           OutlineLevel=1,
                           Start=datetime.datetime(2026, 1, 1),
                           Finish=AMANHA, Duration=480,
                           ResourceNames="Alice",
                           Predecessors="" if i == 1 else "%dFS" % (i - 1))
               for i in range(1, N + 1)]
    proj_g = ProjetoFalso(grandes, nome="Grande")
    instalar_falsos(proj_g)
    antes = contador_g[0]
    await chamar("validate_schedule")
    por_tarefa = (contador_g[0] - antes) / float(N)
    ok("validate_schedule le menos de 20 propriedades por tarefa em %d tarefas" % N,
       por_tarefa < 20, "leu %.1f" % por_tarefa)
    instalar_falsos(proj)
    # --- o que uma leitura COM que falha derruba -------------------------
    #
    # _safe transforma falha em default. Isso e certo onde o default ja
    # significa "ausente", e errado onde ele seria uma AFIRMACAO -- e a
    # diferenca so existe enquanto alguem a defender. Este bloco e a defesa:
    # protecao demais publica numero inventado como se fosse medicao;
    # protecao de menos derruba a varredura inteira por uma tarefa ruim.
    print("\n=== LEITURA COM QUE FALHA ===")

    PROTEGIDOS = {
        "WBS": ("wbs", ""), "Milestone": ("milestone", False),
        "Start": ("start", None), "Finish": ("finish", None),
        "Predecessors": ("predecessors", ""),
        "ResourceNames": ("resource_names", ""), "Notes": ("notes", ""),
        "Critical": ("critical", False), "Active": ("active", True),
        "Text1": ("text1", ""), "Text2": ("text2", ""), "Text3": ("text3", ""),
        "Flag1": ("flag1", False), "Flag2": ("flag2", False),
        "Deadline": ("deadline", None), "Priority": ("priority", 500),
        "TotalSlack": ("total_slack_days", 0), "FreeSlack": ("free_slack_days", 0),
    }
    for propriedade, (campo, esperado) in sorted(PROTEGIDOS.items()):
        ruim = TarefaQueFalha([0], falham=[propriedade], UniqueID=1, ID=1,
                              Name="Ruim", Duration=480)
        try:
            lido = msp_core.task_to_dict(ruim, MPD)
            ok("%s ilegivel nao derruba, vira %r" % (propriedade, esperado),
               lido[campo] == esperado, "veio %r" % (lido.get(campo),))
        except Exception as exc:
            ok("%s ilegivel nao derruba" % propriedade, False,
               "%s: %s" % (type(exc).__name__, exc))

    # E os que devem derrubar. Identidade inventada publica a tarefa errada;
    # estrutura inventada remonta a arvore WBS; medicao inventada vira achado
    # do validate_schedule. Falhar alto e a resposta certa nos tres casos.
    for propriedade in ("UniqueID", "ID", "Name", "OutlineLevel", "Summary",
                        "PercentComplete", "Duration"):
        ruim = TarefaQueFalha([0], falham=[propriedade], UniqueID=1, ID=1,
                              Name="Ruim", Duration=480)
        try:
            msp_core.task_to_dict(ruim, MPD)
            ok("%s ilegivel deve derrubar, nao virar default" % propriedade, False,
               "passou calado")
        except Exception:
            ok("%s ilegivel derruba em vez de inventar" % propriedade, True)

    await payload(ok)


async def payload(ok):
    """O corte de payload: limite, campos vazios e serializacao.

    Medido no cronograma real: get_tasks devolvia 8,07 MB. Estes casos travam
    as tres decisoes que derrubaram isso, e cada uma pode regredir sozinha e
    calada -- um limite que para de valer, um enxugar que volta a publicar
    vazios, um responder que volta a indentar.
    """
    print("\n=== PAYLOAD ===")

    # enxugar
    cheia = {"unique_id": 7, "id": 3, "name": "A", "notes": "", "text2": "x",
             "percent_complete": 0, "total_slack_days": 0.0, "critical": False,
             "start": None, "duration_days": 0, "priority": 500}
    magra = msp_core.enxugar(cheia)
    ok("enxugar tira vazio, None e False",
       "notes" not in magra and "start" not in magra and "critical" not in magra)
    ok("enxugar mantem quem identifica a linha",
       all(k in magra for k in ("unique_id", "id", "name")))
    ok("enxugar mantem zero que e medicao",
       magra["percent_complete"] == 0 and magra["total_slack_days"] == 0.0
       and magra["duration_days"] == 0)
    vazia = msp_core.enxugar({"unique_id": 1, "id": 1, "name": "", "notes": ""})
    ok("enxugar nunca apaga a identidade, nem vazia", set(vazia) == {"unique_id", "id", "name"})

    # recortar
    fatia, pagina = msp_core.recortar(list(range(1000)), limite=200)
    ok("recortar corta e diz que cortou",
       len(fatia) == 200 and pagina["total"] == 1000 and pagina.get("truncated"))
    ok("recortar diz como pegar o resto", "how_to_get_the_rest" in pagina)
    fatia, pagina = msp_core.recortar(list(range(1000)), limite=200, offset=900)
    ok("ultima pagina nao se declara truncada",
       len(fatia) == 100 and not pagina.get("truncated"))
    fatia, pagina = msp_core.recortar(list(range(1000)), limite=-1)
    ok("limite -1 devolve tudo", len(fatia) == 1000 and not pagina.get("truncated"))
    fatia, pagina = msp_core.recortar([1, 2, 3])
    ok("lista menor que o limite sai inteira e sem aviso",
       fatia == [1, 2, 3] and not pagina.get("truncated"))

    # responder
    grande = msp_core.responder({"tasks": [{"x": i} for i in range(2000)]})
    ok("resposta grande sai compacta", "\n" not in grande)
    ok("resposta pequena sai indentada", "\n" in msp_core.responder({"a": 1}))

    # ponta a ponta: um projeto grande o bastante para o limite valer
    contador = [0]
    muitas = [TarefaFalsa(contador, UniqueID=i, ID=i, Name="Tarefa %d" % i,
                          OutlineLevel=2,
                          Start=datetime.datetime(2026, 1, 5),
                          Finish=AMANHA, Duration=480)
              for i in range(1, 501)]
    instalar_falsos(ProjetoFalso(muitas))
    msp_fast.varredura = lambda p: (None, {"backend": "com", "reason": "test"})
    r = await chamar("get_tasks")
    ok("get_tasks para no limite e informa o total",
       len(r["tasks"]) == msp_core.LIMITE_PADRAO and r["count"] == 500
       and r["page"].get("truncated"))
    # `v in (None, "", False)` diria que 0 e vazio, porque 0 == False em
    # Python -- e reprovaria exatamente os zeros que devem ficar.
    def nada(v):
        return v is None or v == "" or v is False
    ok("nenhuma tarefa publicada carrega campo vazio",
       not any(nada(v) for t in r["tasks"] for v in t.values()),
       str([{k: v for k, v in t.items() if nada(v)} for t in r["tasks"][:2]]))
    r2 = await chamar("filter_tasks", filters_json=json.dumps({"limit": -1}))
    ok("filter_tasks com limit -1 devolve tudo", len(r2["tasks"]) == 500)
    # As definicoes das tools sao pagas por toda sessao, antes de qualquer
    # chamada. O "title" que o pydantic gera para cada propriedade custava
    # 8 KB dos 50,7 KB e nao diz nada que o nome do parametro ja nao diga.
    definicoes = await mcp.list_tools()
    esquemas = json.dumps([t.inputSchema for t in definicoes])
    ok("nenhum schema publicado carrega 'title'", '"title"' not in esquemas)

    r3 = await chamar("filter_tasks", filters_json=json.dumps({"offset": 480}))
    ok("filter_tasks pagina a partir do offset",
       len(r3["tasks"]) == 20 and r3["tasks"][0]["unique_id"] == 481)


def main():
    asyncio.run(rodar())
    print("\n" + "=" * 52)
    print("  PASS: %d   FAIL: %d" % (PASS, FAIL))
    print("=" * 52)
    sys.exit(0 if FAIL == 0 else 1)


if __name__ == "__main__":
    main()
