# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

VNStock Agent - MCP server and CLI for Vietnamese stock market data. Wraps the `vnstock` Python library to provide AI tool integration via Model Context Protocol.

Two entry points, one shared data layer:
- `vnstock-mcp` → `vnstock_agent.server:run` - FastMCP server (stdio/SSE/HTTP), 21 tools
- `vnstock-agent` → `vnstock_agent.cli:main` - Click CLI, 22 commands

## Commands

```bash
# Install (editable, with dev deps)
pip install -e ".[dev]"

# Run CLI
VNSTOCK_API_KEY=xxx vnstock-agent history VNM
VNSTOCK_API_KEY=xxx vnstock-agent --format json history VNM

# Run MCP server
VNSTOCK_API_KEY=xxx vnstock-mcp                              # stdio (default)
VNSTOCK_MCP_TRANSPORT=sse vnstock-mcp                         # SSE
VNSTOCK_MCP_TRANSPORT=http vnstock-mcp                        # streamable HTTP
vnstock-agent serve --transport sse --port 8000               # equivalent via CLI

# Lint
ruff check src/

# Test
pytest
```

Note: `pyproject.toml` sets `testpaths = ["tests"]` but no `tests/` directory exists yet in this repo — `pytest` currently collects nothing. `docs/code-standards.md` documents the intended test layout/conventions (`tests/test_core.py`, `tests/test_cli.py`) for when tests are added.

Docker (SSE server on port 8000):
```bash
docker compose up --build
# or: docker build -t vnstock-agent . && docker run -p 8000:8000 -e VNSTOCK_API_KEY=xxx vnstock-agent
```
Dockerfile is `python:3.12-alpine` (musl) — chosen for K8s compatibility; build deps (gcc, musl-dev, python3-dev) are installed and removed in the same layer.

## Architecture

```
src/vnstock_agent/
├── config.py   # env-var settings + ensure_api_key() (stdlib only, no other internal imports)
├── core.py     # shared vnstock wrapper: DataFrame/Series/dict → list[dict]
├── server.py   # FastMCP tool wrappers around core.py (imports core, config)
└── cli.py      # Click command wrappers around core.py (imports core, config)
```

Dependency direction is strictly acyclic: `config → (nothing internal)`, `core → config`, `server → core, config`, `cli → core, config`. `server.py` and `cli.py` never import each other.

Every stock function exists in **three** places doing the same thing at three layers — when adding a new data source/endpoint, touch all three:
1. `core.py` — business logic function, returns `list[dict]`, wraps the vnstock call in `_safe_call()` or a try/except that returns `[{"error": str(e)}]`
2. `server.py` — `@mcp.tool()` wrapper with the same signature, calls the core function and returns `json.dumps(data, ensure_ascii=False, default=str)`
3. `cli.py` — Click `@main.command()` wrapper, uppercases symbol args, calls the core function, and passes the result to `_output()`

### Data pipeline (`core._df_to_records`)

All vnstock calls return heterogeneous shapes (DataFrame, Series, dict with tuple keys, or a tuple of DataFrames for bid/ask pairs). `_df_to_records()` in `core.py` is the single normalization point: flattens MultiIndex columns (`("listing","symbol")` → `"listing_symbol"`), converts NaN → `None`, converts datetime columns to ISO strings, and always returns `list[dict]`. Use it (or `_safe_call()`, which wraps a call + `_df_to_records`) rather than converting DataFrames ad hoc.

### Key decisions / gotchas

- FastMCP: `FastMCP(...)` takes an `instructions` param, not `description`
- vnstock/vnai print registration/banner messages to stdout on import and on `register_user()` — these must stay suppressed (see `config.ensure_api_key()` and `core._suppress_stdout()`) since stdout is the MCP stdio transport's protocol channel
- `trading.price_board()` and other MultiIndex-column results need the tuple-key flattening in `_df_to_records`
- MSN source (used for `fx_history`, `crypto_history`, `world_index_history`) has an upstream timezone bug in vnstock
- `listing_*` and `trading_*` functions instantiate `vnstock.api.listing.Listing` / `vnstock.api.trading.Trading` directly (source lowercased); quote/company/finance functions go through `Vnstock(...).stock(...)` via `core._get_stock()`

## Additional docs

`docs/` has more detailed (and partly aspirational/future-facing) reference material: `system-architecture.md` (request-flow diagrams, transport architecture), `code-standards.md` (naming, docstring format, commit conventions), `codebase-summary.md`, `deployment-guide.md`. Treat forward-looking sections in these (async core, caching layer, HTTP REST API) as roadmap, not current state.
