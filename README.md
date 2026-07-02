# VNStock Agent

MCP server and CLI for Vietnamese stock market data. Integrates [vnstock](https://github.com/thinh-vu/vnstock) with AI tools via Model Context Protocol (MCP).

## Features

- **21 MCP tools** for AI assistants (Claude, GPT, etc.)
- **22 CLI commands** for terminal use
- **3 MCP transports**: stdio, SSE, Streamable HTTP
- Vietnamese stocks, forex, crypto, world indices, mutual funds
- Historical prices, financials, company info, trading data

## Quick Start

### Install

```bash
pip install vnstock-agent
```

Or from source:

```bash
git clone https://github.com/mrgoonie/vnstock-agent.git
cd vnstock-agent
pip install -e .
```

### Set API Key

```bash
export VNSTOCK_API_KEY=your_api_key_here
```

Get a free API key at [vnstocks.com/login](https://vnstocks.com/login).

## CLI Usage

```bash
# Stock history
vnstock-agent history VNM --start 2025-01-01 --end 2025-03-31

# Company overview
vnstock-agent overview FPT

# Financial statements
vnstock-agent balance-sheet VCB --period quarter
vnstock-agent income VCB --period annual
vnstock-agent cashflow VCB
vnstock-agent ratio VCB

# Market data
vnstock-agent symbols                    # All listed symbols
vnstock-agent group --group VN30         # VN30 stocks
vnstock-agent board "VNM,FPT,ACB,VCB"   # Price board
vnstock-agent industries                 # ICB industries

# Company details
vnstock-agent shareholders FPT
vnstock-agent officers FPT
vnstock-agent news FPT
vnstock-agent events FPT

# Intraday & depth
vnstock-agent intraday VNM --page-size 200

# Global markets
vnstock-agent fx EURUSD --start 2025-01-01
vnstock-agent crypto BTC --start 2025-01-01
vnstock-agent index DJI --start 2025-01-01

# Mutual funds
vnstock-agent funds

# JSON output
vnstock-agent --format json history VNM

# FO Market Scorecard (Polo) — cycle phase & allocation dashboard
vnstock-agent scorecard update data/scorecard/FO_Market_Scorecard_Polo_OneSheet_v5.xlsx \
  --dashboard docs/scorecard/index.html
```

## Market Scorecard (Cycle & Allocation)

`vnstock-agent scorecard update FILE` fetches live VN market data via vnstock,
recomputes the "Market_Scorecard_Polo" workbook's cycle/allocation formulas
in Python, writes the fetched values back into the workbook (formulas are
left untouched so it still recalculates if opened in Excel/Sheets), and can
render a standalone HTML dashboard.

Auto-updated every run: VNIndex spot, 20D/6M avg liquidity, MA50/MA200, RSI14,
VN30-basket breadth (% above MA50/MA200, a proxy for full-market breadth),
BTC 30D ROC, and the USD/VND rate. Everything else (valuation table, macro
series, margin/ETF flow, sentiment/news/FDI) has no vnstock data source and
stays a manual weekly input, per the sheet's own Notes/Frequency columns.

```bash
vnstock-agent scorecard update path/to/scorecard.xlsx \
  --dashboard path/to/dashboard.html \
  [--out path/to/other-file.xlsx] [--skip-breadth]
```

A scheduled GitHub Actions workflow (`.github/workflows/scorecard-update.yml`)
runs this on weekday mornings against `data/scorecard/FO_Market_Scorecard_Polo_OneSheet_v5.xlsx`
and commits the refreshed workbook + `docs/scorecard/index.html`. Set the
`VNSTOCK_API_KEY` repo secret for it to run, and enable GitHub Pages on
`docs/` if you want the dashboard served as a URL.

Note: the tool flags two data-integrity issues it found in the original
workbook (cell `E8` holds text instead of a numeric weight, so the
CyclePsych group score doesn't count toward the Overall Market Score; and
row 27's group label has a trailing space so it's excluded from the Flow
group score) — see the CLI/dashboard output for details rather than a
silent "fix."

## MCP Server

### stdio (for Claude Desktop, Cursor, etc.)

Add to your MCP client config:

```json
{
  "mcpServers": {
    "vnstock": {
      "command": "vnstock-mcp",
      "env": {
        "VNSTOCK_API_KEY": "your_api_key_here"
      }
    }
  }
}
```

### SSE Transport

```bash
VNSTOCK_MCP_TRANSPORT=sse VNSTOCK_MCP_PORT=8000 vnstock-mcp
```

### Streamable HTTP Transport

```bash
VNSTOCK_MCP_TRANSPORT=http VNSTOCK_MCP_PORT=8000 vnstock-mcp
```

Or via CLI:

```bash
vnstock-agent serve --transport sse --port 8000
vnstock-agent serve --transport http --port 8000
```

## MCP Tools

| Tool | Description |
|------|-------------|
| `stock_history` | Historical OHLCV price data |
| `stock_intraday` | Today's intraday trading data |
| `stock_price_depth` | Order book / bid-ask data |
| `company_overview` | Company overview info |
| `company_shareholders` | Major shareholders |
| `company_officers` | Management team |
| `company_news` | Company news |
| `company_events` | Dividends, AGM, etc. |
| `financial_balance_sheet` | Balance sheet |
| `financial_income_statement` | Income statement |
| `financial_cash_flow` | Cash flow statement |
| `financial_ratio` | P/E, P/B, ROE, ROA, etc. |
| `listing_all_symbols` | All listed symbols |
| `listing_symbols_by_group` | VN30, HNX30, etc. |
| `listing_symbols_by_exchange` | By exchange (HOSE, HNX, UPCOM) |
| `listing_industries` | ICB industry classification |
| `trading_price_board` | Real-time price board |
| `fx_history` | Forex historical data |
| `crypto_history` | Cryptocurrency data |
| `world_index_history` | World market indices |
| `fund_listing` | Mutual fund listing |

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `VNSTOCK_API_KEY` | (none) | VNStock API key |
| `VNSTOCK_SOURCE` | `VCI` | Default data source (VCI, KBS) |
| `VNSTOCK_MCP_TRANSPORT` | `stdio` | MCP transport (stdio, sse, http) |
| `VNSTOCK_MCP_HOST` | `0.0.0.0` | Server host |
| `VNSTOCK_MCP_PORT` | `8000` | Server port |

## Development

```bash
git clone https://github.com/mrgoonie/vnstock-agent.git
cd vnstock-agent
python -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"

# Run tests
pytest

# Lint
ruff check src/
```

## License

MIT
