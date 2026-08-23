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
    ]


def sem_origem(valor):
    """The response with every 'source' block stripped, at any depth."""
    if isinstance(valor, dict):
        return {k: sem_origem(v) for k, v in valor.items() if k != "source"}
    if isinstance(valor, list):
        return [sem_origem(v) for v in valor]
    return valor


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


def main():
    asyncio.run(rodar())
    print("\n" + "=" * 52)
    print("  PASS: %d   FAIL: %d" % (PASS, FAIL))
    print("=" * 52)
    sys.exit(0 if FAIL == 0 else 1)


if __name__ == "__main__":
    main()
