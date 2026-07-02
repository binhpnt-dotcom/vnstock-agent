"""CLI for VNStock Agent - command-line access to Vietnamese stock market data."""

import json
import sys

import click

from vnstock_agent import core
from vnstock_agent.config import DEFAULT_SOURCE
from vnstock_agent import scorecard as _scorecard


def _output(data, output_format: str):
    """Output data in requested format."""
    if output_format == "json":
        click.echo(json.dumps(data, ensure_ascii=False, indent=2, default=str))
    else:
        # Table format using tabulate
        if not data:
            click.echo("No data returned.")
            return
        try:
            from tabulate import tabulate
            click.echo(tabulate(data, headers="keys", tablefmt="simple"))
        except ImportError:
            click.echo(json.dumps(data, ensure_ascii=False, indent=2, default=str))


@click.group()
@click.version_option(package_name="vnstock-agent")
@click.option("--format", "output_format", type=click.Choice(["table", "json"]), default="table", help="Output format")
@click.pass_context
def main(ctx, output_format):
    """VNStock Agent CLI - Vietnamese stock market data at your fingertips."""
    ctx.ensure_object(dict)
    ctx.obj["format"] = output_format


# --- Quote commands ---


@main.command()
@click.argument("symbol")
@click.option("--start", "-s", default=None, help="Start date (YYYY-MM-DD)")
@click.option("--end", "-e", default=None, help="End date (YYYY-MM-DD)")
@click.option("--interval", "-i", default="1D", help="Interval: 1m,5m,15m,30m,1H,1D,1W,1M")
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def history(ctx, symbol, start, end, interval, source):
    """Get historical OHLCV price data for a stock."""
    data = core.stock_history(symbol.upper(), start, end, interval, source)
    _output(data, ctx.obj["format"])


@main.command()
@click.argument("symbol")
@click.option("--page-size", "-n", default=100, help="Number of records")
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def intraday(ctx, symbol, page_size, source):
    """Get intraday trading data for a stock."""
    data = core.stock_intraday(symbol.upper(), page_size, source)
    _output(data, ctx.obj["format"])


@main.command()
@click.argument("symbol")
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def depth(ctx, symbol, source):
    """Get order book / price depth for a stock."""
    data = core.stock_price_depth(symbol.upper(), source)
    _output(data, ctx.obj["format"])


# --- Company commands ---


@main.command()
@click.argument("symbol")
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def overview(ctx, symbol, source):
    """Get company overview information."""
    data = core.company_overview(symbol.upper(), source)
    _output(data, ctx.obj["format"])


@main.command()
@click.argument("symbol")
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def shareholders(ctx, symbol, source):
    """Get major shareholders of a company."""
    data = core.company_shareholders(symbol.upper(), source)
    _output(data, ctx.obj["format"])


@main.command()
@click.argument("symbol")
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def officers(ctx, symbol, source):
    """Get company officers / management team."""
    data = core.company_officers(symbol.upper(), source)
    _output(data, ctx.obj["format"])


@main.command()
@click.argument("symbol")
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def news(ctx, symbol, source):
    """Get latest company news."""
    data = core.company_news(symbol.upper(), source)
    _output(data, ctx.obj["format"])


@main.command()
@click.argument("symbol")
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def events(ctx, symbol, source):
    """Get company events (dividends, AGM, etc)."""
    data = core.company_events(symbol.upper(), source)
    _output(data, ctx.obj["format"])


# --- Financial commands ---


@main.command()
@click.argument("symbol")
@click.option("--period", "-p", default="quarter", type=click.Choice(["quarter", "annual"]))
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def balance_sheet(ctx, symbol, period, source):
    """Get balance sheet data."""
    data = core.financial_balance_sheet(symbol.upper(), period, source)
    _output(data, ctx.obj["format"])


@main.command()
@click.argument("symbol")
@click.option("--period", "-p", default="quarter", type=click.Choice(["quarter", "annual"]))
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def income(ctx, symbol, period, source):
    """Get income statement data."""
    data = core.financial_income_statement(symbol.upper(), period, source)
    _output(data, ctx.obj["format"])


@main.command()
@click.argument("symbol")
@click.option("--period", "-p", default="quarter", type=click.Choice(["quarter", "annual"]))
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def cashflow(ctx, symbol, period, source):
    """Get cash flow statement data."""
    data = core.financial_cash_flow(symbol.upper(), period, source)
    _output(data, ctx.obj["format"])


@main.command()
@click.argument("symbol")
@click.option("--period", "-p", default="quarter", type=click.Choice(["quarter", "annual"]))
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def ratio(ctx, symbol, period, source):
    """Get financial ratios (P/E, P/B, ROE, etc)."""
    data = core.financial_ratio(symbol.upper(), period, source)
    _output(data, ctx.obj["format"])


# --- Listing commands ---


@main.command()
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def symbols(ctx, source):
    """Get all listed stock symbols."""
    data = core.listing_all_symbols(source)
    _output(data, ctx.obj["format"])


@main.command()
@click.option("--group", "-g", default="VN30", help="Market group: VN30, HNX30, HOSE, etc")
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def group(ctx, group, source):
    """Get symbols in a market group."""
    data = core.listing_symbols_by_group(group, source)
    _output(data, ctx.obj["format"])


@main.command()
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def exchanges(ctx, source):
    """Get symbols grouped by exchange."""
    data = core.listing_symbols_by_exchange(source)
    _output(data, ctx.obj["format"])


@main.command()
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def industries(ctx, source):
    """Get ICB industry classification."""
    data = core.listing_industries(source)
    _output(data, ctx.obj["format"])


# --- Trading commands ---


@main.command()
@click.argument("symbols")
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.pass_context
def board(ctx, symbols, source):
    """Get price board for symbols (comma-separated, e.g. VNM,FPT,ACB)."""
    symbols_list = [s.strip().upper() for s in symbols.split(",") if s.strip()]
    data = core.trading_price_board(symbols_list, source)
    _output(data, ctx.obj["format"])


# --- Global market commands ---


@main.command()
@click.argument("symbol", default="EURUSD")
@click.option("--start", "-s", default=None, help="Start date (YYYY-MM-DD)")
@click.option("--end", "-e", default=None, help="End date (YYYY-MM-DD)")
@click.option("--interval", "-i", default="1D", help="Interval: 1D, 1W, 1M")
@click.pass_context
def fx(ctx, symbol, start, end, interval):
    """Get forex historical price data."""
    data = core.fx_history(symbol.upper(), start, end, interval)
    _output(data, ctx.obj["format"])


@main.command()
@click.argument("symbol", default="BTC")
@click.option("--start", "-s", default=None, help="Start date (YYYY-MM-DD)")
@click.option("--end", "-e", default=None, help="End date (YYYY-MM-DD)")
@click.option("--interval", "-i", default="1D", help="Interval: 1D, 1W, 1M")
@click.pass_context
def crypto(ctx, symbol, start, end, interval):
    """Get cryptocurrency historical price data."""
    data = core.crypto_history(symbol.upper(), start, end, interval)
    _output(data, ctx.obj["format"])


@main.command()
@click.argument("symbol", default="DJI")
@click.option("--start", "-s", default=None, help="Start date (YYYY-MM-DD)")
@click.option("--end", "-e", default=None, help="End date (YYYY-MM-DD)")
@click.option("--interval", "-i", default="1D", help="Interval: 1D, 1W, 1M")
@click.pass_context
def index(ctx, symbol, start, end, interval):
    """Get world market index historical data."""
    data = core.world_index_history(symbol.upper(), start, end, interval)
    _output(data, ctx.obj["format"])


# --- Fund commands ---


@main.command()
@click.pass_context
def funds(ctx):
    """Get list of open-ended mutual funds."""
    data = core.fund_listing()
    _output(data, ctx.obj["format"])


# --- Scorecard commands ---


@main.group()
def scorecard():
    """FO Market Scorecard (Polo) automation: cycle phase & allocation."""


@scorecard.command("update")
@click.argument("file", type=click.Path(exists=True, dir_okay=False))
@click.option("--out", "-o", default=None, help="Write updated workbook to a different path (default: overwrite FILE)")
@click.option("--dashboard", "-d", default=None, help="Path to write an HTML dashboard summary")
@click.option("--source", default=DEFAULT_SOURCE, help="Data source: VCI or KBS")
@click.option("--skip-breadth", is_flag=True, help="Skip VN30 market-breadth fetch (faster, ~30 fewer API calls)")
def scorecard_update(file, out, dashboard, source, skip_breadth):
    """Fetch live VN market data and recompute cycle/allocation in the scorecard workbook."""
    result, auto, applied = _scorecard.update_workbook(
        file, out_path=out, source=source, include_breadth=not skip_breadth
    )
    click.echo(f"Overall Market Score: {result.overall_score:.1f}  ({result.stance})")
    click.echo(f"Cycle Phase: {result.cycle_phase}  (sentiment={result.sentiment_regime}, macro={result.macro_regime})")
    click.echo(f"Suggested Equity Weight: {result.suggested_equity_weight:.1%}  (leverage: {result.leverage_note})")
    click.echo(f"Auto-updated fields: {', '.join(applied) if applied else 'none'}")
    if auto.errors:
        click.echo("Fetch issues:")
        for k, v in auto.errors.items():
            click.echo(f"  - {k}: {v}")
    if result.warnings:
        stale = [w for w in result.warnings if w not in auto.errors.values()]
        if stale:
            click.echo("Detected issues in the source workbook:")
            for w in stale:
                click.echo(f"  - {w}")
    if dashboard:
        _scorecard.write_dashboard(result, auto, applied, dashboard)
        click.echo(f"Dashboard written to {dashboard}")


@scorecard.command("checklist")
@click.argument("file", type=click.Path(exists=True, dir_okay=False))
def scorecard_checklist(file):
    """List which indicators need manual weekly/monthly/quarterly updates (no network calls)."""
    _, ws = _scorecard.load_workbook_sheet(file)
    base = _scorecard.read_base_inputs(ws)
    result = _scorecard.compute(ws, base, _scorecard.AutoData())
    checklist = _scorecard.manual_update_checklist(result.rows)
    for freq in ("Weekly", "Monthly", "Quarterly"):
        items = checklist.get(freq, [])
        click.echo(f"\n{freq} ({len(items)}):")
        for r in items:
            click.echo(f"  - [{r.group.strip()}] {r.factor}")
    other = checklist.get("Other", [])
    if other:
        click.echo(f"\nOther ({len(other)}):")
        for r in other:
            click.echo(f"  - [{r.group.strip()}] {r.factor}")
    auto_count = sum(1 for r in result.rows if r.source == "auto")
    manual_count = sum(len(v) for v in checklist.values())
    click.echo(f"\n{auto_count} indicators are auto-updated by 'scorecard update'; {manual_count} need manual input.")


# --- MCP server command ---


@main.command()
@click.option("--transport", "-t", default="stdio", type=click.Choice(["stdio", "sse", "http"]),
              help="MCP transport mode")
@click.option("--host", default="0.0.0.0", help="Server host (for sse/http)")
@click.option("--port", default=8000, type=int, help="Server port (for sse/http)")
def serve(transport, host, port):
    """Start the MCP server."""
    import os
    os.environ["VNSTOCK_MCP_TRANSPORT"] = transport
    os.environ["VNSTOCK_MCP_HOST"] = host
    os.environ["VNSTOCK_MCP_PORT"] = str(port)
    # Reload config
    from vnstock_agent import server
    server.run()


if __name__ == "__main__":
    main()
