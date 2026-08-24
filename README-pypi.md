# MS Project MCP Server

Controla o Microsoft Project via automação COM através do Model Context Protocol (MCP).

99 ferramentas para ler, editar e analisar cronogramas direto do seu assistente de IA — sem sair do chat.

## Requisitos

| Item | Detalhe |
|------|---------|
| SO | **Windows** (a automação COM não existe em outros sistemas) |
| Microsoft Project | Instalado e **em execução** (testado no MS Project 16.0) |
| Python | 3.10 ou superior |

## Instalação

```bash
pip install msproject-mcp
```

Isso instala as dependências e cria o comando `msproject-mcp`. Teste que ele sobe:

```bash
msproject-mcp
```

Ele fica aguardando no stdin — é o transporte do MCP. `Ctrl+C` encerra.

Caminho rápido de leitura opcional (mpxj + JVM; não disponível em Windows ARM64):

```bash
pip install "msproject-mcp[fast]"
```

> **Antes de usar qualquer ferramenta**: o MS Project precisa estar aberto com um arquivo carregado. O servidor se conecta a uma instância já em execução.

## Registro no cliente MCP

**Claude Code:**

```bash
claude mcp add msproject --scope user -- msproject-mcp
```

**Claude Desktop** — em `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "msproject": {
      "command": "msproject-mcp",
      "args": []
    }
  }
}
```

**OpenAI Codex** — em `~/.codex/config.toml`:

```toml
[mcp_servers.msproject]
command = "msproject-mcp"
args = []
```

## Documentação completa

Inventário das 99 ferramentas, convenções da API, testes e limitações conhecidas: [github.com/devGPL/MS-Procject-MCP](https://github.com/devGPL/MS-Procject-MCP)
