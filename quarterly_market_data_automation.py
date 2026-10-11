#!/usr/bin/env python3
"""Quarterly SIC/FF refresh plus 3-year monthly beta and price-data Excel reports.

Commands:
  refresh            Update SIC->FF48 mapping and Fama/French 3-factor CSV.
  portfolio-report   Fill your Excel template (portfolio_template.xlsx) with fresh betas;
                      save as a new dated file. The template itself is never modified.
  ticker-export       Create one Excel workbook containing monthly data for requested tickers.
  make-template       Create a portfolio-input template workbook.
  factor-refresh      Overwrite the FF5 / Momentum (monthly + daily) / FF48 / ME-breakpoint CSVs.
  monthly-run         Rebuild the SIC mapping only if needed (FF48 definitions changed or a held
                      ticker is unmapped), refresh factor files, then portfolio-report.
  input-run           Rebuild the report when portfolio_input.xlsx content changed (factor files on disk).
  install-schedule    Register launchd jobs: monthly-run on the 1st at 09:00, input-run whenever the input is saved.
  uninstall-schedule  Remove those launchd jobs.
"""
from __future__ import annotations

import argparse
import io
import json
import os
import plistlib
import re
import shutil
import subprocess
import sys
import time
import urllib.request
import zipfile
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd
import yfinance as yf
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

FF_BASE = "https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/ftp/"
FF48_ZIP = FF_BASE + "48_Industry_Portfolios_CSV.zip"
SICCODES_ZIP = FF_BASE + "Siccodes48.zip"  # FF48 SIC-range definitions (unchanged since 2020-01-08)
ME_ZIP = FF_BASE + "ME_Breakpoints_CSV.zip"  # NYSE ME(시가총액) 백분위 브레이크포인트 (Small/Big 분류용)
MONTH_RE = re.compile(r"^\s*\d{6}\s*,")
DAY_RE = re.compile(r"^\s*\d{8}\s*,")

NEW_ROW_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
SIG_POS_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # 연한 초록: 유의한 양(+)
SIG_NEG_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # 연한 빨강: 유의한 음(-)
BOLD = Font(bold=True)

FF5_ZIP = FF_BASE + "F-F_Research_Data_5_Factors_2x3_CSV.zip"
MOM_ZIP = FF_BASE + "F-F_Momentum_Factor_CSV.zip"
FF5_FILE = "F-F_Research_Data_5_Factors_2x3.csv"  # source of Mkt-RF, SMB, HML, RMW, CMA, RF
MOM_FILE = "F-F_Momentum_Factor.csv"
FF5_DAILY_FILE = "F-F_Research_Data_5_Factors_2x3_daily.csv"
MOM_DAILY_FILE = "F-F_Momentum_Factor_daily.csv"
FF48_DAILY_FILE = "48_Industry_Portfolios_daily.csv"
FACTOR_FILES = {  # every French file the report reads; all overwritten by factor-refresh
    FF5_FILE: FF5_ZIP,
    MOM_FILE: MOM_ZIP,
    FF5_DAILY_FILE: FF_BASE + "F-F_Research_Data_5_Factors_2x3_daily_CSV.zip",
    MOM_DAILY_FILE: FF_BASE + "F-F_Momentum_Factor_daily_CSV.zip",
    FF48_DAILY_FILE: FF_BASE + "48_Industry_Portfolios_daily_CSV.zip",
    "48_Industry_Portfolios.csv": FF48_ZIP,
    "ME_Breakpoints.csv": ME_ZIP,
}
STYLE_FACTORS = ["SMB", "HML", "RMW", "CMA", "Mom"]
REGRESSION_FACTORS = ["Mkt-RF"] + STYLE_FACTORS
BETA_LABELS = {
    "Mkt-RF": "Market beta (Mkt-RF)", "SMB": "Size beta (SMB)", "HML": "Value beta (HML)",
    "RMW": "Profitability beta (RMW)", "CMA": "Investment beta (CMA)", "Mom": "Momentum beta (Mom)",
}
REFERENCE_SHEET = "참고용(부풀려짐)"  # one-factor monthly betas, kept only for comparison
TRADING_DAYS = 252
DIMSON_LAGS = 1    # lagged factor days added to the daily regression (thin-trading correction)
VIF_WARN = 5.0
CORR_WARN = 0.5
SERIES_COLORS = ["2A78D6", "EB6834", "1BAF7A", "EDA100", "E87BA4", "008300"]  # fixed categorical order
HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
MDD_BLOCK_DAYS = 21  # bootstrap block length (about one trading month)
MDD_SEED = 42
MDD_TARGET = 0.125      # drawdown budget: portfolio_input.xlsx MDD = total assets x 12.5%
BETA_HALF_LIFE = 126    # trading days; EWMA half-life of the betas used in the MDD simulation
ALPHA_PRIOR = 0.10      # annual spread of true alphas assumed when shrinking estimated alpha
ALPHA_SCENARIOS = ("alpha 0", "shrunk alpha", "historical alpha")
LAUNCH_LABEL = "com.portfolio.monthly-factor-run"
INPUT_WATCH_LABEL = "com.portfolio.input-watch"
QUARTERLY_LABEL = "com.portfolio.quarterly-sic-refresh"  # retired; uninstalled by install-schedule
MAPPING_STATE_FILE = "mapping_state.json"
INPUT_STATE_FILE = "input_state.json"    # content hash of portfolio_input.xlsx the last report used
INPUT_SETTLE_SECONDS = 5  # in the data dir: last FF48 definition date + unmappable tickers


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Quarterly/monthly market-data automation")
    sub = p.add_subparsers(dest="command", required=True)

    r = sub.add_parser("refresh", help="Overwrite current SIC mapping and all factor CSVs")
    r.add_argument("--email", required=True, help="Contact email passed to SEC User-Agent")
    r.add_argument("--ff48-file", default="Siccodes48.txt", help="FF48 SIC definition file")
    r.add_argument("--builder", default="make_us_listed_sic_ff48.py", help="Your supplied SIC-builder script")
    r.add_argument("--workdir", default="data", help="Data directory")
    r.add_argument("--raw-dir", default=str(Path.home() / "PortfolioData" / "sic_builder_output"),
                   help="Where the SIC builder downloads SEC bulk files (~1.5GB); kept outside the iCloud-synced Desktop")
    r.add_argument("--all-exchanges", action="store_true", help="Pass through to SIC builder")

    b = sub.add_parser("portfolio-report", help="Fill your template workbook with portfolio betas")
    b.add_argument("--input", default="portfolio_input.xlsx", help="Excel input, sheet 'Portfolio'")
    b.add_argument("--mapping", default="data/us_listed_sic_ff48.csv", help="Current SIC/FF48 mapping CSV")
    b.add_argument("--template", default="portfolio_template.xlsx", help="Your pre-formatted template workbook")
    b.add_argument("--output-dir", default="reports")
    b.add_argument("--years", type=int, default=3)
    b.add_argument("--market-ticker", default="SPY", help="Used only as fallback if FF market factor is absent")
    b.add_argument("--header-row", type=int, default=1, help="Row number of the header in each template sheet")
    b.add_argument("--alpha", type=float, default=0.05, help="Significance level for green/red factor-correlation cells")
    b.add_argument("--mdd-horizon-days", type=int, default=TRADING_DAYS, help="Expected-MDD horizon in trading days")
    b.add_argument("--mdd-paths", type=int, default=5000, help="Simulated paths for expected MDD")
    b.add_argument("--mdd-target", type=float, default=MDD_TARGET, help="MDD budget (0.125 = 12.5%%) for the recommended weights")
    b.add_argument("--beta-half-life", type=int, default=BETA_HALF_LIFE, help="EWMA half-life (trading days) of simulation betas")
    b.add_argument("--alpha-prior", type=float, default=ALPHA_PRIOR, help="Annual prior spread of alpha used for shrinkage")

    t = sub.add_parser("ticker-export", help="Create one workbook of 3-year monthly ticker data")
    t.add_argument("--tickers", nargs="+", required=True, help="Yahoo Finance tickers, e.g. AAPL MSFT 005930.KS")
    t.add_argument("--output-dir", default="reports")
    t.add_argument("--years", type=int, default=3)

    m = sub.add_parser("make-template", help="Create a portfolio input template")
    m.add_argument("--output", default="portfolio_input.xlsx")

    f = sub.add_parser("factor-refresh", help="Overwrite FF5 / Momentum (monthly + daily) / FF48 / ME-breakpoint CSVs")
    f.add_argument("--workdir", default="data", help="Data directory")

    mr = sub.add_parser("monthly-run", help="Monthly job: SIC check + factor refresh + portfolio-report")
    ir = sub.add_parser("input-run", help="Rebuild the report after portfolio_input.xlsx changes (no factor download)")
    for job in (mr, ir):
        job.add_argument("--workdir", default="data", help="Data directory")
        job.add_argument("--years", type=int, default=3)
        job.add_argument("--input", default="portfolio_input.xlsx", help="Checked for tickers missing from the SIC mapping")
        job.add_argument("--ff48-file", default="Siccodes48.txt", help="FF48 SIC definition file (updated if French changes it)")
        job.add_argument("--email", help="SEC contact email; without it the SIC mapping is never rebuilt automatically")

    s = sub.add_parser("install-schedule", help="Register monthly-run (1st of month) and input-run (on input save) with launchd")
    s.add_argument("--day", type=int, default=1, help="Day of month for monthly-run")
    s.add_argument("--hour", type=int, default=9)
    s.add_argument("--minute", type=int, default=0)
    s.add_argument("--email", help="SEC contact email so monthly-run can rebuild the SIC mapping when needed")

    sub.add_parser("uninstall-schedule", help="Remove the launchd jobs")
    return p


def download(url: str, destination: Path, user_agent: str = "Mozilla/5.0") -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    tmp = destination.with_suffix(destination.suffix + ".part")
    request = urllib.request.Request(url, headers={"User-Agent": user_agent})
    try:
        with urllib.request.urlopen(request, timeout=180) as response, open(tmp, "wb") as f:
            shutil.copyfileobj(response, f)
        tmp.replace(destination)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise


def download_french_csv(zip_url: str, destination: Path) -> None:
    """Download a French Library ZIP and save its enclosed CSV under an exact local name."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    request = urllib.request.Request(zip_url, headers={"User-Agent": "quarterly-research-automation/1.0"})
    with urllib.request.urlopen(request, timeout=120) as response:
        raw = response.read()
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        csv_names = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        if len(csv_names) != 1:
            raise RuntimeError(f"Expected one CSV in {zip_url}; found {csv_names}")
        content = zf.read(csv_names[0])
    tmp = destination.with_suffix(destination.suffix + ".part")
    tmp.write_bytes(content)
    tmp.replace(destination)


def quarterly_refresh(args: argparse.Namespace) -> None:
    """Run the supplied SEC implementation, then replace the canonical current CSV."""
    builder = Path(args.builder)
    definition = Path(args.ff48_file)
    if not builder.exists():
        raise FileNotFoundError(f"SIC builder not found: {builder}")
    if not definition.exists():
        raise FileNotFoundError(f"FF48 definition not found: {definition}")

    workdir = Path(args.workdir)
    run_dir = Path(args.raw_dir).expanduser()
    cmd = [sys.executable, str(builder), "--email", args.email, "--ff48-file", str(definition), "--output-dir", str(run_dir)]
    if args.all_exchanges:
        cmd.append("--all-exchanges")
    subprocess.run(cmd, check=True)

    snapshot_files = sorted(run_dir.glob("us_listed_sic_ff48_*.csv"), key=lambda p: p.stat().st_mtime)
    if not snapshot_files:
        raise RuntimeError("The SIC builder finished but did not produce us_listed_sic_ff48_*.csv")
    current_mapping = workdir / "us_listed_sic_ff48.csv"
    shutil.copy2(snapshot_files[-1], current_mapping)  # canonical file is deliberately overwritten
    print(f"Overwritten SIC mapping: {current_mapping}")
    refresh_factor_files(workdir)


def refresh_factor_files(workdir: Path) -> None:
    """Overwrite every French Library file the report reads. Run monthly so the factor,
    industry and ME-breakpoint series all end on the same (latest published) month."""
    for filename, url in FACTOR_FILES.items():
        path = workdir / filename
        download_french_csv(url, path)
        reader = read_french_daily if "daily" in filename else read_french_breakpoints
        last = reader(path).index.max()
        print(f"Overwritten {path} (data through {last:%Y-%m-%d})")


def factor_refresh(args: argparse.Namespace) -> None:
    refresh_factor_files(Path(args.workdir))


def read_french_section(path: Path, row_re: re.Pattern, date_format: str) -> pd.DataFrame:
    """Read the first data block (monthly YYYYMM or daily YYYYMMDD rows) of a French
    Library CSV, using the line just above it as the header. Values are left unscaled."""
    lines = path.read_text(encoding="latin-1").splitlines()
    data_start = next((i for i, line in enumerate(lines) if row_re.match(line)), None)
    if data_start is None:
        raise ValueError(f"Data section not found in {path}")
    header = data_start - 1
    data_end = next((i for i in range(data_start, len(lines)) if not row_re.match(lines[i])), len(lines))
    df = pd.read_csv(io.StringIO("\n".join(lines[header:data_end])), index_col=0)
    df.index = pd.to_datetime(df.index.astype(str).str.strip(), format=date_format)
    df.index.name = "Date"
    df.columns = [str(c).strip() for c in df.columns]
    return df.apply(pd.to_numeric, errors="coerce")


def read_french_monthly(path: Path) -> pd.DataFrame:
    """Monthly return section, scaled by /100 because these files store percentages."""
    df = read_french_section(path, MONTH_RE, "%Y%m")
    df.index = df.index + pd.offsets.MonthEnd(0)
    return df.mask(df <= -99.99) / 100.0  # French files mark missing values as -99.99 / -999


def read_french_daily(path: Path) -> pd.DataFrame:
    """Daily return file (YYYYMMDD rows), scaled by /100."""
    df = read_french_section(path, DAY_RE, "%Y%m%d")
    return df.mask(df <= -99.99) / 100.0


def read_french_breakpoints(path: Path) -> pd.DataFrame:
    """Read ME Breakpoints (or similar percentile-breakpoint files). No /100 scaling,
    because these values are $ million market-equity levels, not returns."""
    df = read_french_section(path, MONTH_RE, "%Y%m")
    df.index = df.index + pd.offsets.MonthEnd(0)
    return df


def monthly_prices_and_returns(tickers: list[str], years: int) -> tuple[pd.DataFrame, pd.DataFrame, list[dict], pd.DataFrame, pd.DataFrame]:
    """Month-end prices, monthly returns, download issues, and the daily closes / volumes they came from."""
    end = pd.Timestamp.today(tz="UTC").tz_localize(None).normalize() + pd.Timedelta(days=1)
    start = (end - pd.DateOffset(years=years, months=1)).replace(day=1)
    prices, daily, volume, issues = {}, {}, {}, []
    for ticker in tickers:
        symbol = ticker.upper().strip()
        try:
            raw = yf.Ticker(symbol).history(start=start.date(), end=end.date(), auto_adjust=True, actions=False)
            if raw.empty or "Close" not in raw:
                issues.append({"Ticker": symbol, "Status": "No Yahoo Finance price history returned"})
                continue
            series = raw["Close"].copy()
            series.index = pd.to_datetime(series.index).tz_localize(None).normalize()
            daily[symbol] = series
            volume[symbol] = pd.Series(raw["Volume"].to_numpy(), index=series.index) if "Volume" in raw else None
            prices[symbol] = series.resample("ME").last().dropna()
        except Exception as exc:
            issues.append({"Ticker": symbol, "Status": f"Download error: {exc}"})
    px = pd.DataFrame(prices).sort_index()
    ret = px.pct_change().dropna(how="all")
    volume = {k: v for k, v in volume.items() if v is not None}
    return px, ret, issues, pd.DataFrame(daily).sort_index(), pd.DataFrame(volume).sort_index()


def beta(y: pd.Series, x: pd.Series) -> tuple[float, int]:
    pair = pd.concat([y.rename("y"), x.rename("x")], axis=1, sort=True).dropna()
    if len(pair) < 24 or pair["x"].var(ddof=1) == 0:
        return np.nan, len(pair)
    return pair["y"].cov(pair["x"]) / pair["x"].var(ddof=1), len(pair)


def ols(Y: np.ndarray, X: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    coef, *_ = np.linalg.lstsq(X, Y, rcond=None)
    return coef, Y - X @ coef


def newey_west_cov(X: np.ndarray, resid: np.ndarray, lags: int) -> np.ndarray:
    """Newey-West (Bartlett kernel) HAC covariance matrix of OLS coefficients."""
    n, k = X.shape
    xe = X * resid[:, None]
    meat = xe.T @ xe
    for lag in range(1, lags + 1):
        g = xe[lag:].T @ xe[:-lag]
        meat += (1 - lag / (lags + 1)) * (g + g.T)
    bread = np.linalg.pinv(X.T @ X)
    return bread @ meat @ bread * n / (n - k)


def dimson_design(factors: pd.DataFrame, lags: int = DIMSON_LAGS) -> pd.DataFrame:
    """Each factor plus its 1..lags trading-day lags, ordered factor by factor."""
    return pd.concat(
        {f"{c} L{lag}": factors[c].shift(lag) for c in factors.columns for lag in range(lags + 1)}, axis=1,
    )


def factor_regression_daily(y: pd.Series, factors: pd.DataFrame, lags: int = DIMSON_LAGS, min_obs: int = 200) -> dict:
    """Joint daily OLS of excess return on all factors and their lags (Dimson 1979). A
    factor's beta is the sum of its same-day and lagged coefficients, which removes the
    downward bias stale prices put on thinly traded small caps. Standard errors are
    Newey-West HAC, so heteroskedastic / autocorrelated daily residuals don't overstate
    significance.

    Variance decomposition: y = alpha + sum_f component_f + resid, so
    cov(component_f, y) / var(y) per factor plus var(resid) / var(y) sums exactly to 1.
    A negative share means that factor exposure offsets (hedges) the others."""
    from scipy import stats

    design = dimson_design(factors, lags)
    data = pd.concat([y.rename("__y__"), design], axis=1, sort=True).dropna()
    n = len(data)
    out = {"N": n}
    if n < min_obs:
        return out
    Y = data["__y__"].to_numpy(float)
    X = np.column_stack([np.ones(n), data[design.columns].to_numpy(float)])
    coef, resid = ols(Y, X)
    k = X.shape[1]
    nw_lags = int(np.floor(4 * (n / 100) ** (2 / 9)))
    cov = newey_west_cov(X, resid, nw_lags)
    dof = n - k

    def test(weights: np.ndarray) -> tuple[float, float, float]:
        est = float(weights @ coef)
        t = est / np.sqrt(weights @ cov @ weights)
        return est, t, 2 * stats.t.sf(abs(t), dof)

    var_y = Y.var(ddof=1)
    r2 = 1 - (resid @ resid) / ((Y - Y.mean()) ** 2).sum()
    intercept = np.zeros(k)
    intercept[0] = 1
    alpha, alpha_t, alpha_p = test(intercept)
    alpha_half = stats.t.ppf(0.975, dof) * np.sqrt(cov[0, 0])  # 95% CI half-width (HAC SE), daily units
    out.update({
        "NW lags": nw_lags, "Annualized volatility": np.sqrt(var_y * TRADING_DAYS),
        "Alpha (annualized)": alpha * TRADING_DAYS,
        "Alpha 95% CI low": (alpha - alpha_half) * TRADING_DAYS, "Alpha 95% CI high": (alpha + alpha_half) * TRADING_DAYS,
        "Alpha t": alpha_t, "Alpha p": alpha_p,
        "R2": r2, "Adj R2": 1 - (1 - r2) * (n - 1) / dof,
    })
    factor_share = 0.0
    for i, f in enumerate(factors.columns):
        idx = [1 + i * (lags + 1) + lag for lag in range(lags + 1)]  # matches dimson_design order
        w = np.zeros(k)
        w[idx] = 1
        out[f"{f} beta"], out[f"{f} t"], out[f"{f} p"] = test(w)
        share = np.cov(X[:, idx] @ coef[idx], Y)[0, 1] / var_y
        out[f"{f} risk share"] = share
        factor_share += share
    out["Factor-driven share"] = factor_share
    out["Idiosyncratic risk share"] = resid.var(ddof=1) / var_y
    out["_alpha_daily"], out["_alpha_se_daily"] = alpha, float(np.sqrt(cov[0, 0]))  # for alpha shrinkage
    return out


def rolling_factor_betas(y: pd.Series, factors: pd.DataFrame, window: int = TRADING_DAYS,
                         lags: int = DIMSON_LAGS, min_obs: int = 200) -> pd.DataFrame:
    """Joint Dimson betas re-estimated on the trailing `window` trading days at each month end."""
    design = dimson_design(factors, lags)
    data = pd.concat([y.rename("__y__"), design], axis=1, sort=True).dropna()
    rows = {}
    for d in data.groupby(data.index.to_period("M")).tail(1).index:
        win = data.loc[:d].tail(window)
        if len(win) < min_obs:
            continue
        X = np.column_stack([np.ones(len(win)), win[design.columns].to_numpy(float)])
        coef, _ = ols(win["__y__"].to_numpy(float), X)
        rows[d] = {f: coef[1 + i * (lags + 1): 1 + (i + 1) * (lags + 1)].sum() for i, f in enumerate(factors.columns)}
    return pd.DataFrame.from_dict(rows, orient="index")


def factor_diagnostics(factors: pd.DataFrame) -> pd.DataFrame:
    """Factor correlation matrix plus VIF (diagonal of the inverse correlation matrix):
    how much each factor's beta variance is inflated by overlap with the other factors."""
    corr = factors.corr()
    return corr.assign(VIF=np.diag(np.linalg.inv(corr.to_numpy())))


def risk_contributions(returns: pd.DataFrame, weights: pd.Series) -> pd.Series:
    """Euler risk contribution of each holding (sums to 1) from the sample covariance."""
    w = weights.dropna()
    w = w[w.index.isin(returns.columns)]
    if w.empty or not w.sum():
        return pd.Series(dtype=float)
    w = w / w.sum()
    cov = returns[w.index].dropna().cov()
    port_var = float(w @ cov @ w)
    return w * (cov @ w) / port_var if port_var > 0 else pd.Series(np.nan, index=w.index)


def ewma_factor_model(y: pd.Series, factors: pd.DataFrame, half_life: int = BETA_HALF_LIFE,
                      lags: int = DIMSON_LAGS, min_obs: int = 200) -> dict | None:
    """Current factor exposure for the drawdown simulation: the same Dimson 6-factor
    regression, but weighted so a day `half_life` trading days ago counts half as much as
    today (WLS = OLS on sqrt(weight)-scaled rows). Coefficient covariance is Newey-West on
    the weighted rows, so the simulation can draw betas within their estimation error."""
    design = dimson_design(factors, lags)
    data = pd.concat([y.rename("__y__"), design], axis=1, sort=True).dropna()
    n = len(data)
    if n < min_obs:
        return None
    Y = data["__y__"].to_numpy(float)
    X = np.column_stack([np.ones(n), data[design.columns].to_numpy(float)])
    sw = np.sqrt(0.5 ** ((n - 1 - np.arange(n)) / half_life))
    coef, resid_w = ols(Y * sw, X * sw[:, None])
    cov = newey_west_cov(X * sw[:, None], resid_w, int(np.floor(4 * (n / 100) ** (2 / 9))))
    resid = Y - X @ coef
    lag_groups = [[i * (lags + 1) + lag for lag in range(lags + 1)] for i in range(len(factors.columns))]
    return {
        "dates": data.index,
        "coef": coef[1:], "cov": cov[1:, 1:],                       # factor coefficients (incl. lags)
        "X": pd.DataFrame(X[:, 1:], index=data.index),              # factor design rows
        "resid": pd.Series(resid, index=data.index),
        "groups": lag_groups,
        "betas": {f: coef[1:][g].sum() for f, g in zip(factors.columns, lag_groups)},
    }


def shrink_alpha(alpha_daily: float, se_daily: float, prior_annual: float = ALPHA_PRIOR) -> float:
    """Shrink an estimated alpha toward 0 by its own noise: alpha x tau^2 / (tau^2 + SE^2),
    tau = prior spread of true alphas (default 10%/yr). Noisy alphas shrink almost to 0."""
    tau = prior_annual / TRADING_DAYS
    return alpha_daily * tau ** 2 / (tau ** 2 + se_daily ** 2) if np.isfinite(se_daily) else 0.0


def simulate_wealth_paths(models: dict[str, dict], alphas: dict[str, dict[str, float]], premia: pd.Series,
                          rf_daily: float, horizon: int, n_paths: int, block: int = MDD_BLOCK_DAYS,
                          seed: int = MDD_SEED) -> tuple[dict[str, dict[str, np.ndarray]], dict[str, float], int]:
    """Simulated wealth paths (start = 1) per alpha scenario and name over `horizon` days.

    Daily return on path p, simulated day d (= historical day h):
        RF + B_p . long-run factor premia + alpha(scenario)
        + (factor rows on day h - their window mean) . b_p + (residual on day h - its mean)
    b_p = factor coefficients drawn per path from N(current EWMA estimate, Newey-West
    covariance), B_p = b_p summed over Dimson lags. Historical days are drawn in `block`-day
    runs and the SAME days are used for every name, so cross-stock and cross-factor
    correlation is preserved. Seed fixed so month-to-month changes come from the data."""
    common = None
    for m in models.values():
        common = m["dates"] if common is None else common.intersection(m["dates"])
    if not models or common is None or len(common) < block:
        return {}, {}, 0
    rng = np.random.default_rng(seed)
    n_blocks = -(-horizon // block)
    starts = rng.integers(0, len(common) - block + 1, size=(n_paths, n_blocks))
    days = (starts[:, :, None] + np.arange(block)).reshape(n_paths, -1)[:, :horizon]
    premia_vec = premia[REGRESSION_FACTORS].to_numpy()
    wealth = {scenario: {} for scenario in ALPHA_SCENARIOS}
    expected = {}
    for name, m in models.items():
        X = m["X"].loc[common].to_numpy()
        X = X - X.mean(axis=0)
        resid = m["resid"].loc[common].to_numpy()
        resid = resid - resid.mean()
        draws = rng.multivariate_normal(m["coef"], m["cov"], size=n_paths, method="eigh")
        shock = resid[days]
        for k in range(X.shape[1]):
            shock += X[:, k][days] * draws[:, k][:, None]
        summed = np.column_stack([draws[:, g].sum(axis=1) for g in m["groups"]])  # per-path factor betas
        mu = rf_daily + summed @ premia_vec
        expected[name] = float(mu.mean())
        for scenario in ALPHA_SCENARIOS:
            r = mu[:, None] + alphas.get(name, {}).get(scenario, 0.0) + shock
            wealth[scenario][name] = np.cumprod(1 + np.clip(r, -0.99, None), axis=1)
    return wealth, expected, len(common)


def path_mdd(wealth: np.ndarray) -> np.ndarray:
    """Maximum drawdown of each path (row), measured from the starting value 1."""
    peak = np.maximum(np.maximum.accumulate(wealth, axis=1), 1.0)
    return np.minimum((wealth / peak - 1).min(axis=1), 0.0)


def path_max_loss(wealth: np.ndarray) -> np.ndarray:
    """Largest loss below the starting value (cost) on each path: min(wealth) - 1, capped at 0.
    Unlike MDD, gains made and later given back are not counted as a loss."""
    return np.minimum(wealth.min(axis=1) - 1, 0.0)


def loss_stats(loss: np.ndarray, target: float) -> dict:
    return {
        "Expected max loss vs cost (mean)": loss.mean(),
        "Worst 5% max loss": np.percentile(loss, 5),
        f"P(max loss <= -{target:.1%})": (loss <= -target).mean(),
    }


def mdd_stats(mdd: np.ndarray, target: float) -> dict:
    return {
        "Expected MDD (mean)": mdd.mean(),
        "Median MDD": np.median(mdd),
        "Worst 5% MDD": np.percentile(mdd, 5),
        f"P(MDD <= -{target:.1%})": (mdd <= -target).mean(),
        "P(MDD <= -20%)": (mdd <= -0.20).mean(),
        "P(MDD <= -50%)": (mdd <= -0.50).mean(),
    }


def buy_and_hold(wealth: dict[str, np.ndarray], weights: pd.Series) -> tuple[np.ndarray | None, pd.Series]:
    """Portfolio wealth paths when the weights are bought at the start and held."""
    w = weights.dropna()
    w = w[w.index.isin(list(wealth))]
    if w.empty or not w.sum():
        return None, w
    w = w / w.sum()
    return sum(w[t] * wealth[t] for t in w.index), w


def exposure_for_target_loss(port_paths: np.ndarray, rf_daily: float, target: float) -> float:
    """Share k of capital in the (current-mix) stock portfolio, rest in cash earning RF, at
    which the expected (mean) max loss vs cost equals -target. The loss grows with k, so
    bisection on [0, 1]; returns 1.0 if a fully invested portfolio stays within the target."""
    cash = (1 + rf_daily) ** np.arange(1, port_paths.shape[1] + 1)

    def mean_mdd(k: float) -> float:
        return path_max_loss(k * port_paths + (1 - k) * cash).mean()

    if mean_mdd(1.0) >= -target:
        return 1.0
    lo, hi = 0.0, 1.0
    for _ in range(40):
        mid = (lo + hi) / 2
        lo, hi = (mid, hi) if mean_mdd(mid) >= -target else (lo, mid)
    return lo


def historical_mdd(prices: pd.Series) -> float:
    p = prices.dropna()
    return float((p / p.cummax() - 1).min()) if len(p) > 1 else np.nan


def build_mdd_table(wealth: dict[str, dict], expected: dict, models: dict, alphas: dict, weights: pd.Series,
                    navs: pd.Series | None, window_px: pd.DataFrame, horizon: int, target: float) -> pd.DataFrame:
    base = wealth.get("alpha 0", {})
    rows = []
    for t, w_paths in base.items():
        if t not in weights.index:
            continue  # Test-only names are simulated for Test_Comparison but not listed here
        stats = mdd_stats(path_mdd(w_paths), target)
        losses = loss_stats(path_max_loss(w_paths), target)
        nav = navs.get(t) if navs is not None else None
        rows.append({
            "Ticker": t, "Weight": weights.get(t),
            "Annualized volatility": models[t].get("Annualized volatility"),
            "Expected return (horizon)": (1 + expected[t]) ** horizon - 1,
            **losses,
            "Expected max loss ($)": losses["Expected max loss vs cost (mean)"] * nav if nav is not None and pd.notna(nav) else None,
            **stats,
            "Expected MDD - shrunk alpha": path_mdd(wealth["shrunk alpha"][t]).mean(),
            "Expected MDD - historical alpha": path_mdd(wealth["historical alpha"][t]).mean(),
            "Alpha used - shrunk (annualized)": alphas[t]["shrunk alpha"] * TRADING_DAYS,
            "Alpha used - historical (annualized)": alphas[t]["historical alpha"] * TRADING_DAYS,
            "Historical MDD (lookback)": historical_mdd(window_px[t]) if t in window_px else np.nan,
            "Expected MDD ($)": stats["Expected MDD (mean)"] * nav if nav is not None and pd.notna(nav) else None,
        })
    port_paths, w = buy_and_hold(base, weights)
    if port_paths is None:
        return pd.DataFrame(rows)
    stats = mdd_stats(path_mdd(port_paths), target)
    losses = loss_stats(path_max_loss(port_paths), target)
    held_px = window_px[w.index].dropna()
    total_nav = navs.reindex(w.index).sum() if navs is not None and navs.reindex(w.index).notna().all() else None
    rows.append({
        "Ticker": "PORTFOLIO", "Weight": 1.0,
        "Annualized volatility": models["PORTFOLIO"].get("Annualized volatility") if "PORTFOLIO" in models else None,
        "Expected return (horizon)": port_paths[:, -1].mean() - 1,
        **losses,
        "Expected max loss ($)": losses["Expected max loss vs cost (mean)"] * total_nav if total_nav else None,
        **stats,
        "Expected MDD - shrunk alpha": path_mdd(buy_and_hold(wealth["shrunk alpha"], weights)[0]).mean(),
        "Expected MDD - historical alpha": path_mdd(buy_and_hold(wealth["historical alpha"], weights)[0]).mean(),
        "Historical MDD (lookback)": historical_mdd((held_px / held_px.iloc[0]).mul(w, axis=1).sum(axis=1)) if len(held_px) else np.nan,
        "Expected MDD ($)": stats["Expected MDD (mean)"] * total_nav if total_nav else None,
    })
    stock_rows = pd.DataFrame(rows[:-1]).set_index("Ticker")
    rows.append({
        "Ticker": "Weighted avg of holdings",
        "Expected max loss vs cost (mean)": (stock_rows["Expected max loss vs cost (mean)"] * w).sum(),
        "Expected MDD (mean)": (stock_rows["Expected MDD (mean)"] * w).sum(),
        "Worst 5% MDD": (stock_rows["Worst 5% MDD"] * w).sum(),
    })
    return pd.DataFrame(rows)


def resolve_weights(sheet: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    """Clean the Ticker column; Weight = NAV / total NAV when every row has a NAV,
    otherwise the sheet's Weight column is used as typed."""
    df = sheet.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if "Ticker" not in df.columns:
        raise ValueError("Portfolio/Test sheet must contain a 'Ticker' column.")
    df = df[df["Ticker"].notna()]
    df["Ticker"] = df["Ticker"].astype(str).str.upper().str.strip()
    df = df[df["Ticker"].ne("")].drop_duplicates("Ticker")
    nav = pd.to_numeric(df["NAV"], errors="coerce") if "NAV" in df.columns else None
    if nav is not None and len(df) and nav.notna().all() and nav.sum() > 0:
        df["Weight"] = nav / nav.sum()
        return df, "NAV / total NAV"
    df["Weight"] = pd.to_numeric(df["Weight"], errors="coerce") if "Weight" in df.columns else np.nan
    return df, "Weight column"


def ensure_factor_file(data_dir: Path, filename: str) -> Path:
    """Download a French Library file only if it is missing; factor-refresh keeps them current."""
    path = data_dir / filename
    if not path.exists():
        download_french_csv(FACTOR_FILES[filename], path)
    return path


def approx_market_equity(tickers: list[str], prices: pd.DataFrame) -> pd.DataFrame:
    """Approximate monthly market equity ($mm) = month-end price * current shares outstanding.
    Caveat: shares outstanding is Yahoo's *current* figure applied retroactively, so tickers
    with major buybacks/issuance/splits in the lookback window will have some drift."""
    me = {}
    for ticker in tickers:
        if ticker not in prices.columns:
            continue
        shares = None
        try:
            shares = yf.Ticker(ticker).fast_info.get("shares")
        except Exception:
            shares = None
        if not shares:
            try:
                shares = yf.Ticker(ticker).info.get("sharesOutstanding")
            except Exception:
                shares = None
        if not shares:
            continue
        me[ticker] = prices[ticker] * shares / 1_000_000  # match Ken French's $mm scale
    return pd.DataFrame(me)


def classify_size(me: pd.Series, breakpoints_50: pd.Series) -> str:
    """Classify Small/Big using the most recent overlapping month vs. NYSE median ME."""
    pair = pd.concat([me.rename("me"), breakpoints_50.rename("bp")], axis=1, sort=True).dropna()
    if pair.empty:
        return "N/A"
    latest = pair.iloc[-1]
    return "Big" if latest["me"] >= latest["bp"] else "Small"


# ---------------------------------------------------------------------------
# Template-fill helpers: open a pre-formatted workbook, drop data in, save new file.
# ---------------------------------------------------------------------------

def clear_sheet_data(ws, header_row: int, n_cols: int, max_row: int = 2000) -> None:
    """Blank out any leftover values below the header from a previous run."""
    for r in range(header_row + 1, max_row + 1):
        for c in range(1, n_cols + 1):
            ws.cell(row=r, column=c).value = None  # ws.cell(..., value=None) would leave the old value


def write_df_to_sheet(ws, df: pd.DataFrame, header_row: int, include_index: bool = False) -> None:
    """Write a DataFrame's values starting one row below header_row, column A (or B if
    include_index=True). Header labels are rewritten from the DataFrame so added columns
    (e.g. new factor betas) and changed tickers stay in sync with the template; a header
    cell whose label changes takes the template's first header style and a fitting width."""
    start_col = 2 if include_index else 1
    n_cols = start_col - 1 + len(df.columns)
    old_max_col = ws.max_column
    clear_sheet_data(ws, header_row, max(n_cols, old_max_col))
    header_style = ws.cell(row=header_row, column=1)
    for c, label in enumerate(df.columns, start=start_col):
        cell = ws.cell(row=header_row, column=c)
        if cell.value != label:
            cell.value = label
            cell._style = header_style._style
            ws.column_dimensions[cell.column_letter].width = max(len(str(label)) + 4, 12)
    for c in range(n_cols + 1, old_max_col + 1):
        ws.cell(row=header_row, column=c).value = None
    for r, (idx, row) in enumerate(df.iterrows()):
        excel_row = header_row + 1 + r
        if include_index:
            ws.cell(row=excel_row, column=1, value=idx if not isinstance(idx, pd.Timestamp) else idx.date())
        for c, value in enumerate(row):
            if isinstance(value, (pd.Timestamp,)):
                value = value.date()
            elif isinstance(value, float) and np.isnan(value):
                value = None
            ws.cell(row=excel_row, column=start_col + c, value=value)


def fill_template(template_path: Path, output_path: Path, header_row: int, sheets: dict[str, tuple[pd.DataFrame, bool]]) -> None:
    """sheets: {sheet_name: (dataframe, include_index)}"""
    wb = load_workbook(template_path)
    for sheet_name, (df, include_index) in sheets.items():
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Template is missing expected sheet '{sheet_name}'. Found: {wb.sheetnames}")
        write_df_to_sheet(wb[sheet_name], df, header_row, include_index)
    wb.save(output_path)


def create_portfolio_report(args: argparse.Namespace) -> Path:
    """Every beta, volatility and risk number in the report comes from DAILY excess returns
    over the lookback window. Monthly data only feeds Monthly_Returns / Monthly_Prices and
    the one-factor reference sheet."""
    inp = Path(args.input)
    mapping_path = Path(args.mapping)
    template_path = Path(args.template)
    if not inp.exists():
        raise FileNotFoundError(f"Input workbook not found: {inp}. Run make-template first.")
    if not mapping_path.exists():
        raise FileNotFoundError(f"SIC mapping not found: {mapping_path}. Run refresh first.")
    if not template_path.exists():
        raise FileNotFoundError(f"Template workbook not found: {template_path}. Create it once with your desired formatting.")

    input_sheets = pd.read_excel(inp, sheet_name=None)
    if "Portfolio" not in input_sheets:
        raise ValueError(f"{inp} must contain a 'Portfolio' sheet.")
    portfolio, weight_source = resolve_weights(input_sheets["Portfolio"])
    test, test_source = None, None
    if "Test" in input_sheets:
        test, test_source = resolve_weights(input_sheets["Test"])
        if test["Weight"].notna().sum() == 0:
            test = None
    print(f"Portfolio: {dict(zip(portfolio['Ticker'], portfolio['Weight'].round(4)))} ({weight_source})")
    if test is not None:
        print(f"Test:      {dict(zip(test['Ticker'], test['Weight'].round(4)))} ({test_source})")

    mapping = pd.read_csv(mapping_path, dtype={"ticker": str})
    mapping["ticker"] = mapping["ticker"].str.upper().str.strip()
    mapping = mapping.drop_duplicates("ticker", keep="last")
    meta = portfolio.merge(
        mapping[["ticker", "sic", "sic_description", "ff48_code", "ff48_short", "ff48_name"]],
        how="left", left_on="Ticker", right_on="ticker",
    ).drop(columns="ticker")

    tickers = meta["Ticker"].tolist()
    test_tickers = test["Ticker"].tolist() if test is not None else []
    all_prices, all_returns, issues, daily_px, daily_vol = monthly_prices_and_returns(
        list(dict.fromkeys(tickers + test_tickers)), args.years,
    )
    held = [t for t in tickers if t in all_prices.columns]
    prices, returns = all_prices[held], all_returns[held].dropna(how="all")
    data_dir = mapping_path.parent

    # --- Daily factor data (FF5 2x3 + Momentum + FF48 industries) over the lookback window ---
    daily_all = read_french_daily(ensure_factor_file(data_dir, FF5_DAILY_FILE)).join(
        read_french_daily(ensure_factor_file(data_dir, MOM_DAILY_FILE))[["Mom"]], how="inner",
    ).dropna()
    daily_end = daily_all.index.max()
    daily_factors = daily_all.loc[daily_end - pd.DateOffset(years=args.years) + pd.Timedelta(days=1):]
    calendar = daily_factors.index
    daily_rf = daily_factors["RF"]
    market = daily_factors["Mkt-RF"]
    F = daily_factors[REGRESSION_FACTORS]
    industry_daily = read_french_daily(ensure_factor_file(data_dir, FF48_DAILY_FILE)).reindex(calendar)
    daily_ret = daily_px.pct_change().reindex(calendar)
    daily_ex = daily_ret.sub(daily_rf, axis=0)

    def portfolio_excess(weights: pd.Series) -> pd.Series:
        w = weights.dropna()
        w = w[w.index.isin(daily_ret.columns)]
        w = w / w.sum()
        return daily_ret[w.index].mul(w, axis=1).sum(axis=1, min_count=len(w)) - daily_rf

    current_w = portfolio.set_index("Ticker")["Weight"]
    model_inputs = {t: daily_ex[t] for t in held}
    model_inputs["PORTFOLIO"] = portfolio_excess(current_w)
    daily_results = {name: factor_regression_daily(y, F) for name, y in model_inputs.items()}
    rolling = {name: rolling_factor_betas(y, F) for name, y in model_inputs.items()}
    port_model = daily_results["PORTFOLIO"]

    # --- Size classification (latest month vs NYSE median ME) ---
    me_breakpoints = read_french_breakpoints(ensure_factor_file(data_dir, "ME_Breakpoints.csv"))
    median_col = "50" if "50" in me_breakpoints.columns else me_breakpoints.columns[len(me_breakpoints.columns) // 2]
    median_bp = me_breakpoints[median_col]
    market_equity = approx_market_equity(tickers, prices)

    asset_rows = []
    for _, row in meta.iterrows():
        ticker, ind = row["Ticker"], row["ff48_short"]
        res = daily_results.get(ticker, {})
        ind_beta = np.nan
        if ticker in daily_ex and pd.notna(ind) and ind in industry_daily.columns:
            # Industry exposure net of the market: y ~ Mkt-RF + (industry - market), daily Dimson
            ind_model = factor_regression_daily(daily_ex[ticker], pd.DataFrame({
                "Mkt-RF": market, "Industry-Mkt": industry_daily[ind] - daily_rf - market,
            }))
            ind_beta = ind_model.get("Industry-Mkt beta", np.nan)
        size_label = classify_size(
            market_equity[ticker] if ticker in market_equity else pd.Series(dtype=float), median_bp,
        )
        asset_rows.append({
            "Ticker": ticker,
            "Weight": row["Weight"],
            **{label: res.get(f"{f} beta", np.nan) for f, label in BETA_LABELS.items()},
            "Industry beta (ex-market)": ind_beta,
            "Size classification (latest)": size_label,
            "SIC description": row["sic_description"],
            "FF48 industry": ind,
        })
    assets = pd.DataFrame(asset_rows)

    industry_rows = []
    for ind in sorted(industry_daily.columns):
        res = factor_regression_daily(industry_daily[ind] - daily_rf, F[["Mkt-RF"]])
        industry_rows.append({
            "FF48 Industry": ind,
            "Industry beta to market (daily Dimson)": res.get("Mkt-RF beta", np.nan),
            "Observations (trading days)": res["N"],
        })
    industry_df = pd.DataFrame(industry_rows)

    # --- One-factor MONTHLY betas: kept only as a labelled reference (they are inflated) ---
    ff5_m = read_french_monthly(ensure_factor_file(data_dir, FF5_FILE))
    mom_m = read_french_monthly(ensure_factor_file(data_dir, MOM_FILE))
    factors_m = ff5_m.join(mom_m[["Mom"]], how="inner")
    one_factor = one_factor_monthly_betas(returns, factors_m, current_w)

    # --- Daily volatility / covariance / Euler risk contribution (one basis for all) ---
    risk_contrib = risk_contributions(daily_ex, current_w)
    risk_window = daily_ex[risk_contrib.index].dropna()

    # --- Expected MDD: simulate every held (and Test) name on the same bootstrapped days ---
    regs = {t: daily_results[t] for t in held}
    for t in test_tickers:
        if t not in regs and t in daily_ex.columns:
            regs[t] = factor_regression_daily(daily_ex[t], F)
    sim_models = {}
    for t in regs:
        m = ewma_factor_model(daily_ex[t], F, args.beta_half_life)
        if m is not None:
            sim_models[t] = m
    alphas = {t: {
        "alpha 0": 0.0,
        "shrunk alpha": shrink_alpha(regs[t].get("_alpha_daily", 0.0), regs[t].get("_alpha_se_daily", np.inf), args.alpha_prior),
        "historical alpha": regs[t].get("_alpha_daily", 0.0),
    } for t in sim_models}
    rf_now = daily_rf.tail(21).mean()
    wealth, expected, sim_days = simulate_wealth_paths(
        sim_models, alphas, daily_all[REGRESSION_FACTORS].mean(), rf_now, args.mdd_horizon_days, args.mdd_paths,
    )
    navs = pd.to_numeric(portfolio.set_index("Ticker")["NAV"], errors="coerce") if "NAV" in portfolio.columns else None
    window_px = daily_px.loc[calendar.min():daily_end]
    mdd_table = build_mdd_table(wealth, expected, {**regs, "PORTFOLIO": port_model}, alphas, current_w, navs,
                                window_px, args.mdd_horizon_days, args.mdd_target)
    port_mdd = mdd_table.set_index("Ticker").loc["PORTFOLIO"] if "PORTFOLIO" in set(mdd_table.get("Ticker", [])) else pd.Series(dtype=object)
    port_paths, port_w = buy_and_hold(wealth.get("alpha 0", {}), current_w)
    recommended = []
    if port_paths is not None:
        exposure = exposure_for_target_loss(port_paths, rf_now, args.mdd_target)
        recommended = [(t, exposure * port_w[t]) for t in port_w.index] + [("현금 (Cash)", 1 - exposure)]

    summary = pd.DataFrame(
        [
            ("Report date", date.today().isoformat()),
            ("Lookback years", args.years),
            ("Weight source", weight_source),
            *[(f"Portfolio {f} beta (joint 6-factor, daily Dimson)", port_model.get(f"{f} beta", np.nan)) for f in REGRESSION_FACTORS],
            ("Portfolio annualized volatility (daily)", port_model.get("Annualized volatility", np.nan)),
            ("Portfolio variance share - factors (daily)", port_model.get("Factor-driven share", np.nan)),
            ("Portfolio variance share - idiosyncratic (daily)", port_model.get("Idiosyncratic risk share", np.nan)),
            *[(f"Risk contribution - {t} (daily)", v) for t, v in risk_contrib.items()],
            (f"Portfolio expected max loss vs cost ({args.mdd_horizon_days}d, simulated mean)", port_mdd.get("Expected max loss vs cost (mean)", np.nan)),
            (f"Portfolio worst-5% max loss vs cost ({args.mdd_horizon_days}d)", port_mdd.get("Worst 5% max loss", np.nan)),
            ("Portfolio expected max loss ($, total NAV)", port_mdd.get("Expected max loss ($)", np.nan)),
            (f"Portfolio expected MDD from peak ({args.mdd_horizon_days}d, simulated mean)", port_mdd.get("Expected MDD (mean)", np.nan)),
            (f"Portfolio worst-5% MDD ({args.mdd_horizon_days}d)", port_mdd.get("Worst 5% MDD", np.nan)),
            ("Portfolio expected MDD ($, total NAV)", port_mdd.get("Expected MDD ($)", np.nan)),
            ("Return convention", "Daily Yahoo adjusted-close simple return minus RF for every beta / volatility / risk number; "
                                  "Monthly_* sheets and the reference sheet are monthly"),
            ("Volatility / risk method", f"Std and covariance of daily excess returns x sqrt({TRADING_DAYS}); "
                                         "risk contribution = Euler share of portfolio variance (sums to 100%)"),
            ("Factor model", "Joint Mkt-RF, SMB, HML, RMW, CMA (FF5 2x3 daily) + Mom; Dimson 1-day lag; Newey-West SEs"),
            ("Industry benchmark", "FF48 daily value-weighted industry; Industry beta = coefficient on (industry - market) "
                                   "with the market also in the regression"),
            ("Size classification", "Small/Big vs NYSE median ME (approx, current shares outstanding)"),
            ("Daily factor data through", f"{daily_end:%Y-%m-%d}"),
            ("Daily window used", f"{calendar.min():%Y-%m-%d} ~ {daily_end:%Y-%m-%d} ({len(calendar)} trading days)"),
        ],
        columns=["Metric", "Value"],
    )

    quality = issues or [{"Ticker": "", "Status": "All downloads succeeded"}]
    quality_df = pd.DataFrame(quality + daily_data_checks(daily_px, daily_vol, calendar))

    Path(args.output_dir).mkdir(parents=True, exist_ok=True)
    output = Path(args.output_dir) / f"portfolio_beta_report_{date.today():%Y%m%d}.xlsx"

    fill_template(
        template_path=template_path,
        output_path=output,
        header_row=args.header_row,
        sheets={
            "Summary": (summary, False),
            "Asset_Betas": (assets, False),
            "Industry_Betas": (industry_df, False),
            "Monthly_Returns": (returns, True),  # plain monthly returns; risk stats live on Daily_Risk
            "Monthly_Prices": (prices, True),
            "Data_Quality": (quality_df, False),
        },
    )

    comparison = None
    if test is not None:
        test_w = test.set_index("Ticker")["Weight"]
        test_model = factor_regression_daily(portfolio_excess(test_w), F)
        comparison = build_test_comparison(
            current_w, test_w, port_model, test_model,
            risk_contributions(daily_ex, current_w), risk_contributions(daily_ex, test_w),
            portfolio_mdd(wealth.get("alpha 0", {}), current_w, args.mdd_target),
            portfolio_mdd(wealth.get("alpha 0", {}), test_w, args.mdd_target),
        )

    wb = load_workbook(output)
    format_summary(wb["Summary"], args.header_row)
    write_recommended_weights(wb["Summary"], recommended, args.mdd_target)
    format_asset_betas(wb["Asset_Betas"], args.header_row)
    regression_rows = [{"Ticker": n, **{k: v for k, v in r.items() if not k.startswith("_")}} for n, r in daily_results.items()]
    after = "Asset_Betas"
    after = write_regression_sheet(wb, after, pd.DataFrame(regression_rows), args.alpha)
    after = write_factor_risk_sheet(wb, after, daily_results)
    after = write_daily_risk_sheet(wb, after, current_w, risk_contrib, risk_window, port_model)
    after = write_mdd_sheet(wb, after, mdd_table, sim_models, args, sim_days)
    after = write_diagnostics_sheet(wb, after, factor_diagnostics(F))
    after = write_rolling_sheet(wb, after, rolling)
    if comparison is not None:
        after = write_test_sheet(wb, after, comparison, args.alpha, test_source)
    write_reference_sheet(wb, wb.sheetnames[-1], one_factor)
    wb.save(output)

    print(f"Created: {output}")
    return output


def one_factor_monthly_betas(returns: pd.DataFrame, factors_m: pd.DataFrame, weights: pd.Series) -> pd.DataFrame:
    """Each factor regressed ALONE on monthly excess returns (the original method). Kept for
    reference only: a factor on its own also absorbs the market and the other factors it
    co-moves with, so these betas are inflated relative to the joint daily model."""
    rows = []
    for t in returns.columns:
        y = returns[t] - factors_m["RF"]
        row = {"Ticker": t, "Weight": weights.get(t)}
        for f, label in BETA_LABELS.items():
            row[label], row["Months used"] = beta(y, factors_m[f])
        rows.append(row)
    df = pd.DataFrame(rows)
    w = df.set_index("Ticker")["Weight"].dropna()
    if not df.empty and w.sum():
        w = w / w.sum()
        weighted = {"Ticker": "PORTFOLIO (weighted)", "Weight": 1.0}
        for label in BETA_LABELS.values():
            weighted[label] = (df.set_index("Ticker")[label].reindex(w.index) * w).sum()
        df = pd.concat([df, pd.DataFrame([weighted])], ignore_index=True)
    return df


def daily_data_checks(daily_px: pd.DataFrame, daily_vol: pd.DataFrame, calendar: pd.DatetimeIndex) -> list[dict]:
    """Per-ticker check of the daily series used by the model: gaps against the factor
    trading calendar, zero-volume days, and the longest run of unchanged closes (a run of
    identical prices on zero volume would point to a trading halt or filled-in data)."""
    rows = []
    for t in daily_px.columns:
        px = daily_px[t].reindex(calendar)
        first = px.first_valid_index()
        if first is None:
            rows.append({"Ticker": t, "Status": "Daily check: no prices inside the factor window"})
            continue
        px = px.loc[first:]
        vol = daily_vol[t].reindex(px.index) if t in daily_vol else pd.Series(np.nan, index=px.index)
        unchanged = px.diff().eq(0)
        longest = int(unchanged.groupby((~unchanged).cumsum()).sum().max()) if unchanged.any() else 0
        zero_vol_unchanged = int((unchanged & vol.eq(0)).sum())
        rows.append({"Ticker": t, "Status": (
            f"Daily check {first:%Y-%m-%d}~{px.index[-1]:%Y-%m-%d}: {int(px.notna().sum())} trading days, "
            f"{int(px.isna().sum())} missing vs factor calendar, {int(vol.eq(0).sum())} zero-volume days, "
            f"longest unchanged-close run {longest} day(s), {zero_vol_unchanged} unchanged closes on zero volume"
        )})
    return rows


# ---------------------------------------------------------------------------
# Analysis sheets added after the template is filled (not part of the template).
# ---------------------------------------------------------------------------

def new_sheet(wb, name: str, after: str):
    if name in wb.sheetnames:
        del wb[name]
    return wb.create_sheet(name, index=wb.sheetnames.index(after) + 1)


def write_header(ws, columns: list[str], row: int = 1, min_width: int = 10) -> None:
    for c, label in enumerate(columns, start=1):
        cell = ws.cell(row=row, column=c, value=label)
        cell.font, cell.fill = BOLD, HEADER_FILL
        ws.column_dimensions[cell.column_letter].width = max(len(str(label)) + 3, min_width)


def cell_value(value):
    if isinstance(value, (float, np.floating)) and np.isnan(value):
        return None
    if isinstance(value, np.generic):
        return value.item()
    return value


def write_notes(ws, start_row: int, notes: list[str]) -> None:
    for i, text in enumerate(notes):
        ws.cell(row=start_row + i, column=1, value=text)


def shade_significant(cell, coef, p, alpha: float) -> None:
    if coef is not None and p is not None and p < alpha:
        cell.fill = SIG_POS_FILL if coef > 0 else SIG_NEG_FILL


def write_regression_sheet(wb, after: str, reg: pd.DataFrame, alpha: float) -> str:
    """Factor_Regression: joint daily 6-factor betas with t / p. Beta (and alpha) cells are
    light green when significantly positive and light red when significantly negative."""
    ws = new_sheet(wb, "Factor_Regression", after)
    columns = ["Ticker", "N", "NW lags", "Alpha (annualized)", "Alpha 95% CI low", "Alpha 95% CI high",
               "Alpha t", "Alpha p", "R2", "Adj R2"]
    for f in REGRESSION_FACTORS:
        columns += [f"{f} beta", f"{f} t", f"{f} p"]
    reg = reg.reindex(columns=columns)
    write_header(ws, columns)
    for r, row in enumerate(reg.itertuples(index=False), start=2):
        values = dict(zip(columns, row))
        for c, label in enumerate(columns, start=1):
            cell = ws.cell(row=r, column=c, value=cell_value(values[label]))
            if label not in ("Ticker", "N", "NW lags"):
                cell.number_format = "0.000"
        for coef_label, p_label in [("Alpha (annualized)", "Alpha p")] + [(f"{f} beta", f"{f} p") for f in REGRESSION_FACTORS]:
            shade_significant(ws.cell(row=r, column=columns.index(coef_label) + 1),
                              cell_value(values[coef_label]), cell_value(values[p_label]), alpha)
        if values["Ticker"] == "PORTFOLIO":
            ws.cell(row=r, column=1).font = BOLD
    write_notes(ws, len(reg) + 3, [
        f"Daily excess return (Yahoo daily return - RF) regressed jointly on {', '.join(REGRESSION_FACTORS)} "
        f"(FF5 2x3 daily + daily Momentum) and their {DIMSON_LAGS}-day lags, with intercept (OLS).",
        "Beta = same-day + lagged coefficient (Dimson): corrects betas biased toward 0 by stale prices in thinly traded stocks.",
        "t / p use Newey-West HAC standard errors (NW lags column = Bartlett lag length). N = trading days used.",
        f"Green = coefficient significantly > 0, red = significantly < 0 (two-sided p < {alpha}). Alpha is annualized (x{TRADING_DAYS}).",
        "Alpha 95% CI = alpha +/- t(0.975, N - parameters) x Newey-West standard error, annualized. "
        "If the interval contains 0, the alpha is not distinguishable from zero.",
        "SMB beta > 0: moves with small caps over large caps. HML > 0: value over growth. RMW > 0: robust over weak profitability. "
        "CMA > 0: conservative over aggressive investment. Mom > 0: past winners over past losers.",
        "PORTFOLIO = report weights (rescaled to sum to 1) applied to the ticker daily returns.",
    ])
    ws.freeze_panes = "B2"
    return ws.title


def write_factor_risk_sheet(wb, after: str, results: dict[str, dict]) -> str:
    """Factor_Risk: share of each name's daily return variance explained by each factor."""
    ws = new_sheet(wb, "Factor_Risk", after)
    share_cols = [f"{f} risk share" for f in REGRESSION_FACTORS] + ["Idiosyncratic risk share", "Factor-driven share"]
    columns = ["Ticker", "Annualized volatility", "R2"] + share_cols
    write_header(ws, columns, min_width=14)
    for r, (name, res) in enumerate(results.items(), start=2):
        ws.cell(row=r, column=1, value=name).font = BOLD if name == "PORTFOLIO" else Font()
        for c, label in enumerate(columns[1:], start=2):
            cell = ws.cell(row=r, column=c, value=cell_value(res.get(label, np.nan)))
            cell.number_format = "0.000" if label == "R2" else "0.0%"
    write_notes(ws, len(results) + 3, [
        "Variance decomposition of daily excess returns from the joint 6-factor Dimson model on Factor_Regression.",
        "y = alpha + sum of factor components + residual, so each factor share = cov(factor component, y) / var(y), "
        "idiosyncratic share = var(residual) / var(y); factor shares + idiosyncratic share = 100%.",
        "Factor-driven share = sum of the six factor shares (= R2). A negative factor share means that exposure offsets "
        "(hedges) risk coming from the other factors.",
        "Annualized volatility = std of daily excess return x sqrt(252).",
    ])
    ws.freeze_panes = "B2"
    return ws.title


def write_diagnostics_sheet(wb, after: str, diag: pd.DataFrame) -> str:
    """Factor_Diagnostics: factor correlation matrix + VIF over the regression window."""
    ws = new_sheet(wb, "Factor_Diagnostics", after)
    columns = ["Factor"] + list(diag.columns)
    write_header(ws, columns, min_width=10)
    for r, (name, row) in enumerate(diag.iterrows(), start=2):
        ws.cell(row=r, column=1, value=name).font = BOLD
        for c, label in enumerate(diag.columns, start=2):
            cell = ws.cell(row=r, column=c, value=cell_value(row[label]))
            cell.number_format = "0.00"
            if label == "VIF" and row[label] > VIF_WARN:
                cell.fill = SIG_NEG_FILL
            elif label != "VIF" and label != name and abs(row[label]) >= CORR_WARN:
                cell.fill = SIG_NEG_FILL
    write_notes(ws, len(diag) + 3, [
        "Daily factor correlations over the same window as Factor_Regression, with each factor's VIF "
        "(variance inflation factor = diagonal of the inverse correlation matrix).",
        f"Red = |correlation| >= {CORR_WARN} or VIF > {VIF_WARN}: those factors overlap, so their joint betas can swing "
        "in opposite directions and should be read together rather than one by one.",
    ])
    return ws.title


def write_rolling_sheet(wb, after: str, rolling: dict[str, pd.DataFrame]) -> str:
    """Rolling_Betas: trailing 1-year joint betas at each month end, one line chart per name."""
    from openpyxl.chart import LineChart, Reference

    ws = new_sheet(wb, "Rolling_Betas", after)
    chart_rows = 17
    top = len(rolling) * chart_rows + 2  # charts first, data table underneath
    dates = sorted(set().union(*[df.index for df in rolling.values()]))
    columns = ["Date"] + [f"{name} {f}" for name in rolling for f in REGRESSION_FACTORS]
    write_header(ws, columns, row=top, min_width=9)
    for r, d in enumerate(dates, start=top + 1):
        ws.cell(row=r, column=1, value=d.date()).number_format = "yyyy-mm-dd"
        c = 2
        for df in rolling.values():
            for f in REGRESSION_FACTORS:
                value = df.at[d, f] if d in df.index and f in df.columns else np.nan
                ws.cell(row=r, column=c, value=cell_value(value)).number_format = "0.00"
                c += 1
    last_row = top + len(dates)
    write_notes(ws, last_row + 2, [
        f"Joint 6-factor Dimson betas re-estimated on the trailing {TRADING_DAYS} trading days at each month end "
        "(points need >= 200 trading days, so the series starts about a year into the lookback).",
    ])
    if not dates:
        return ws.title

    for i, name in enumerate(rolling):
        chart = LineChart()
        chart.title = f"{name}: rolling 1-year joint betas"
        chart.y_axis.title = "Beta"
        chart.x_axis.number_format = "yyyy-mm"
        chart.x_axis.delete = False
        chart.y_axis.delete = False
        chart.legend.position = "b"
        chart.height, chart.width = 8, 22
        first_col = 2 + i * len(REGRESSION_FACTORS)
        data = Reference(ws, min_col=first_col, max_col=first_col + len(REGRESSION_FACTORS) - 1, min_row=top, max_row=last_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=top + 1, max_row=last_row))
        for series, color in zip(chart.series, SERIES_COLORS):
            series.graphicalProperties.line.solidFill = color
            series.graphicalProperties.line.width = 19050  # 1.5pt
            series.smooth = False
            series.marker.symbol = "none"
        ws.add_chart(chart, f"A{1 + i * chart_rows}")
    return ws.title


def write_mdd_sheet(wb, after: str, table: pd.DataFrame, sim_models: dict, args: argparse.Namespace, sim_days: int) -> str:
    """Expected_MDD: simulated max drawdown per holding and for the portfolio, alpha
    scenarios, and the current (EWMA) betas the simulation used."""
    ws = new_sheet(wb, "Expected_MDD", after)
    target = args.mdd_target
    columns = ["Ticker", "Weight", "Annualized volatility", "Expected return (horizon)",
               "Expected max loss vs cost (mean)", "Worst 5% max loss", f"P(max loss <= -{target:.1%})",
               "Expected max loss ($)", "Expected MDD (mean)",
               "Median MDD", "Worst 5% MDD", f"P(MDD <= -{target:.1%})", "P(MDD <= -20%)", "P(MDD <= -50%)",
               "Expected MDD - shrunk alpha", "Expected MDD - historical alpha",
               "Alpha used - shrunk (annualized)", "Alpha used - historical (annualized)",
               "Historical MDD (lookback)", "Expected MDD ($)"]
    table = table.reindex(columns=columns)
    write_header(ws, columns, min_width=13)
    for r, row in enumerate(table.to_dict("records"), start=2):
        for c, label in enumerate(columns, start=1):
            cell = ws.cell(row=r, column=c, value=cell_value(row[label]))
            if label in ("Expected MDD ($)", "Expected max loss ($)"):
                cell.number_format = "#,##0"
            elif label != "Ticker":
                cell.number_format = "0.0%"
        if row["Ticker"] in ("PORTFOLIO", "Weighted avg of holdings"):
            for c in range(1, len(columns) + 1):
                ws.cell(row=r, column=c).font = BOLD
                if row["Ticker"] == "PORTFOLIO":
                    ws.cell(row=r, column=c).fill = NEW_ROW_FILL
    ws.column_dimensions["A"].width = 24

    top = len(table) + 3
    ws.cell(row=top, column=1, value=f"Betas used in the simulation (EWMA, half-life {args.beta_half_life} trading days)").font = BOLD
    write_header(ws, ["Ticker"] + [f"{f} beta" for f in REGRESSION_FACTORS], row=top + 1, min_width=13)
    for r, (name, m) in enumerate(sim_models.items(), start=top + 2):
        ws.cell(row=r, column=1, value=name)
        for c, f in enumerate(REGRESSION_FACTORS, start=2):
            ws.cell(row=r, column=c, value=cell_value(m["betas"][f])).number_format = "0.000"
    write_notes(ws, top + len(sim_models) + 3, [
        f"Max loss vs cost = lowest value over the next {args.mdd_horizon_days} trading days relative to the starting value "
        f"(gains given back are not a loss). This is the measure behind the -{target:.1%} budget and the recommended weights.",
        f"MDD = largest peak-to-trough fall over the same horizon (counts gains given back), {args.mdd_paths:,} simulated paths per name.",
        f"Historical days are re-drawn from the last {sim_days} common trading days in {MDD_BLOCK_DAYS}-day blocks; every name "
        "uses the same days, so stock and factor correlations carry into the portfolio.",
        "Daily return = RF + beta x long-run factor premium (since 1963) + alpha (scenario) + historical factor shock x beta "
        "+ historical residual.",
        f"Betas are the CURRENT exposure (days weighted with a {args.beta_half_life}-day half-life) and are re-drawn on every path "
        "within their Newey-West estimation error, so beta uncertainty widens the MDD distribution.",
        "Main columns use alpha = 0. Scenario columns: shrunk alpha = estimated alpha x tau^2 / (tau^2 + SE^2) with "
        f"tau = {args.alpha_prior:.0%}/yr (noisy alphas shrink toward 0); historical alpha = full estimate (optimistic if it was luck).",
        "PORTFOLIO = current weights bought at the start and held. 'Weighted avg of holdings' = plain weighted average of the "
        "stock MDDs; the gap to PORTFOLIO is the diversification benefit.",
        "Expected MDD ($) = expected MDD x NAV. Historical MDD = realized over the lookback window. "
        f"Recommended weights for a -{target:.1%} expected max loss vs cost are at the bottom of Summary.",
    ])
    ws.freeze_panes = "B2"
    return ws.title


def write_recommended_weights(ws, recommended: list[tuple[str, float]], target: float) -> None:
    """Below the last used Summary row (5 rows down): column A ticker, column B the weight
    that brings the portfolio's expected max loss vs cost to -target (current mix scaled, rest cash)."""
    if not recommended:
        return
    last = max((r for r in range(1, ws.max_row + 1) if any(ws.cell(row=r, column=c).value is not None for c in (1, 2))), default=1)
    top = last + 5
    for c, label in enumerate(["종목", f"추천 비중 (원금 대비 예상 최대손실 -{target:.1%})"], start=1):
        cell = ws.cell(row=top, column=c, value=label)
        cell.font, cell.fill = BOLD, HEADER_FILL
    for r, (name, weight) in enumerate(recommended, start=top + 1):
        ws.cell(row=r, column=1, value=name)
        ws.cell(row=r, column=2, value=cell_value(weight)).number_format = "0.0%"


def format_summary(ws, header_row: int) -> None:
    """Number formats for the Summary values, and the key portfolio numbers in yellow."""
    for r in range(header_row + 1, ws.max_row + 1):
        metric, cell = str(ws.cell(row=r, column=1).value or ""), ws.cell(row=r, column=2)
        if not isinstance(cell.value, (int, float)) or metric == "Lookback years":
            continue
        if "beta" in metric:
            cell.number_format = "0.000"
        elif "($" in metric:
            cell.number_format = "#,##0"
        else:
            cell.number_format = "0.0%"
        if metric.startswith(("Portfolio", "Risk contribution")):
            ws.cell(row=r, column=1).fill = cell.fill = NEW_ROW_FILL
            ws.cell(row=r, column=1).font = BOLD


def format_asset_betas(ws, header_row: int) -> None:
    for c in range(1, ws.max_column + 1):
        label = str(ws.cell(row=header_row, column=c).value or "")
        fmt = "0.0%" if label == "Weight" else "0.000" if "beta" in label else None
        if fmt:
            for r in range(header_row + 1, ws.max_row + 1):
                ws.cell(row=r, column=c).number_format = fmt
    write_notes(ws, ws.max_row + 2, [
        "Betas = joint 6-factor daily model (same numbers as Factor_Regression, Dimson 1-day lag). "
        "Industry beta = daily coefficient on (FF48 industry - market) with the market in the same regression.",
        "One-factor monthly betas (the old method) are on the last sheet, labelled as reference only.",
    ])


def write_daily_risk_sheet(wb, after: str, weights: pd.Series, contrib: pd.Series, window: pd.DataFrame,
                           port_model: dict) -> str:
    """Daily_Risk: the single source for volatility, covariance, correlation and risk contribution."""
    ws = new_sheet(wb, "Daily_Risk", after)
    columns = ["Ticker", "Weight", "Annualized volatility", "Risk contribution"]
    write_header(ws, columns, min_width=14)
    w = weights.reindex(contrib.index)
    w = w / w.sum() if w.sum() else w
    vol = window.std() * np.sqrt(TRADING_DAYS)
    r = 2
    for t in contrib.index:
        for c, value in enumerate([t, w[t], vol[t], contrib[t]], start=1):
            ws.cell(row=r, column=c, value=cell_value(value)).number_format = "0.0%"
        r += 1
    for c, value in enumerate(["PORTFOLIO", 1.0, port_model.get("Annualized volatility"), contrib.sum() if len(contrib) else None], start=1):
        cell = ws.cell(row=r, column=c, value=cell_value(value))
        cell.number_format, cell.font = "0.0%", BOLD

    def matrix(top: int, title: str, m: pd.DataFrame, fmt: str) -> int:
        ws.cell(row=top, column=1, value=title).font = BOLD
        write_header(ws, [""] + list(m.columns), row=top + 1, min_width=14)
        for i, (name, row) in enumerate(m.iterrows(), start=top + 2):
            ws.cell(row=i, column=1, value=name).font = BOLD
            for j, value in enumerate(row, start=2):
                ws.cell(row=i, column=j, value=cell_value(value)).number_format = fmt
        return top + 2 + len(m)

    nxt = matrix(r + 3, "Correlation matrix (daily excess returns)", window.corr(), "0.000")
    nxt = matrix(nxt + 2, f"Covariance matrix (daily excess returns, annualized x{TRADING_DAYS})",
                 window.cov() * TRADING_DAYS, "0.0000")
    write_notes(ws, nxt + 2, [
        f"All numbers use the same {len(window)} daily excess returns (days where every holding traded) over the lookback window.",
        "Risk contribution = Euler share w_i x (Sigma w)_i / (w' Sigma w): sums to 100% of portfolio variance.",
        "Portfolio volatility matches Factor_Risk / Factor_Regression (same daily window and weights).",
    ])
    ws.column_dimensions["A"].width = 14
    return ws.title


def write_reference_sheet(wb, after: str, df: pd.DataFrame) -> str:
    """참고용(부풀려짐): the original one-factor monthly betas, kept only for comparison."""
    ws = new_sheet(wb, REFERENCE_SHEET, after)
    columns = ["Ticker", "Weight"] + list(BETA_LABELS.values()) + ["Months used"]
    df = df.reindex(columns=columns)
    write_header(ws, columns, min_width=12)
    for r, row in enumerate(df.to_dict("records"), start=2):
        for c, label in enumerate(columns, start=1):
            cell = ws.cell(row=r, column=c, value=cell_value(row[label]))
            cell.number_format = "0.0%" if label == "Weight" else "0" if label == "Months used" else "0.000"
        if str(row["Ticker"]).startswith("PORTFOLIO"):
            ws.cell(row=r, column=1).font = BOLD
    ws.column_dimensions["A"].width = 22
    write_notes(ws, len(df) + 3, [
        "참고용(부풀려짐): 예전 방식인 '팩터 하나씩 따로 회귀한 월간 베타'예요. 의사결정에는 Asset_Betas / Factor_Regression 값을 쓰세요.",
        "팩터 하나만 넣으면 그 팩터가 시장과 다른 팩터의 움직임까지 같이 흡수해서 베타가 실제보다 크게 나와요 (누락변수 편의).",
        "Each factor regressed alone on monthly excess returns (FF5 2x3 monthly + Momentum). PORTFOLIO = weight-averaged betas.",
    ])
    return ws.title


def portfolio_mdd(wealth: dict, weights: pd.Series, target: float) -> dict:
    paths, _ = buy_and_hold(wealth, weights)
    return {**loss_stats(path_max_loss(paths), target), **mdd_stats(path_mdd(paths), target)} if paths is not None else {}


def build_test_comparison(current_w: pd.Series, test_w: pd.Series, current: dict, test: dict,
                          current_rc: pd.Series, test_rc: pd.Series,
                          current_mdd: dict, test_mdd: dict) -> pd.DataFrame:
    """Rows of Metric / Current / Test / Change (+ p-values used for shading betas)."""
    def norm(w: pd.Series) -> pd.Series:
        w = w.dropna()
        return w / w.sum() if w.sum() else w

    cw, tw = norm(current_w), norm(test_w)
    rows = []
    for t in dict.fromkeys(list(cw.index) + list(tw.index)):
        rows.append({"Metric": f"Weight - {t}", "Current": cw.get(t, 0.0), "Test": tw.get(t, 0.0), "fmt": "0.0%"})
    rows.append({"Metric": "Annualized volatility (daily)", "Current": current.get("Annualized volatility"),
                 "Test": test.get("Annualized volatility"), "fmt": "0.0%"})
    for label in ["Expected max loss vs cost (mean)", "Worst 5% max loss", "Expected MDD (mean)", "Worst 5% MDD"]:
        rows.append({"Metric": label, "Current": current_mdd.get(label), "Test": test_mdd.get(label), "fmt": "0.0%"})
    rows.append({"Metric": "Alpha (annualized)", "Current": current.get("Alpha (annualized)"),
                 "Test": test.get("Alpha (annualized)"), "fmt": "0.000",
                 "p_current": current.get("Alpha p"), "p_test": test.get("Alpha p")})
    rows.append({"Metric": "R2", "Current": current.get("R2"), "Test": test.get("R2"), "fmt": "0.000"})
    for f in REGRESSION_FACTORS:
        rows.append({"Metric": f"{f} beta", "Current": current.get(f"{f} beta"), "Test": test.get(f"{f} beta"),
                     "fmt": "0.000", "p_current": current.get(f"{f} p"), "p_test": test.get(f"{f} p")})
    for label in [f"{f} risk share" for f in REGRESSION_FACTORS] + ["Idiosyncratic risk share"]:
        rows.append({"Metric": label, "Current": current.get(label), "Test": test.get(label), "fmt": "0.0%"})
    for t in dict.fromkeys(list(current_rc.index) + list(test_rc.index)):
        rows.append({"Metric": f"Risk contribution - {t}", "Current": current_rc.get(t, 0.0),
                     "Test": test_rc.get(t, 0.0), "fmt": "0.0%"})
    df = pd.DataFrame(rows)
    df["Change"] = pd.to_numeric(df["Test"], errors="coerce") - pd.to_numeric(df["Current"], errors="coerce")
    return df


def write_test_sheet(wb, after: str, comparison: pd.DataFrame, alpha: float, test_source: str) -> str:
    """Test_Comparison: current portfolio vs the 'Test' sheet of the input workbook."""
    ws = new_sheet(wb, "Test_Comparison", after)
    write_header(ws, ["Metric", "Current", "Test", "Change"], min_width=14)
    ws.column_dimensions["A"].width = 34
    for r, row in enumerate(comparison.to_dict("records"), start=2):
        ws.cell(row=r, column=1, value=row["Metric"])
        for c, key in enumerate(["Current", "Test", "Change"], start=2):
            cell = ws.cell(row=r, column=c, value=cell_value(row[key]))
            cell.number_format = row["fmt"]
        shade_significant(ws.cell(row=r, column=2), cell_value(row["Current"]), cell_value(row.get("p_current")), alpha)
        shade_significant(ws.cell(row=r, column=3), cell_value(row["Test"]), cell_value(row.get("p_test")), alpha)
    write_notes(ws, len(comparison) + 3, [
        f"Test = the 'Test' sheet of the input workbook (weights from: {test_source}), run through the same daily 6-factor model.",
        "If every Current and Test value matches, the Test sheet holds the same tickers and weights as Portfolio.",
        f"Beta / alpha cells: green = significantly > 0, red = significantly < 0 (p < {alpha}). Change = Test - Current.",
        "Risk contribution = Euler share of daily portfolio variance by holding (sums to 100%).",
    ])
    ws.freeze_panes = "B2"
    return ws.title


def create_ticker_export(args: argparse.Namespace) -> None:
    tickers = list(dict.fromkeys(t.upper().strip() for t in args.tickers if t.strip()))
    prices, returns, issues, _, _ = monthly_prices_and_returns(tickers, args.years)
    long_prices = prices.rename_axis("Date").reset_index().melt("Date", var_name="Ticker", value_name="Adjusted Close")
    long_returns = returns.rename_axis("Date").reset_index().melt("Date", var_name="Ticker", value_name="Monthly Return")
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output = output_dir / f"ticker_monthly_data_{date.today():%Y%m%d}.xlsx"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame({"Ticker": tickers, "Lookback years": args.years, "Data source": "Yahoo Finance via yfinance; adjusted close"}).to_excel(writer, sheet_name="Readme", index=False)
        long_prices.to_excel(writer, sheet_name="Monthly_Prices", index=False)
        long_returns.to_excel(writer, sheet_name="Monthly_Returns", index=False)
        prices.to_excel(writer, sheet_name="Prices_Wide")
        returns.to_excel(writer, sheet_name="Returns_Wide")
        pd.DataFrame(issues or [{"Ticker": "", "Status": "All downloads succeeded"}]).to_excel(writer, sheet_name="Data_Quality", index=False)
    print(f"Created: {output}")


def make_template(args: argparse.Namespace) -> None:
    template = pd.DataFrame({"Ticker": ["AAPL", "MSFT"], "Weight": [0.50, 0.50], "NAV": [None, None]})
    with pd.ExcelWriter(args.output, engine="openpyxl") as writer:
        template.to_excel(writer, sheet_name="Portfolio", index=False)
        template.to_excel(writer, sheet_name="Test", index=False)  # what-if portfolio for Test_Comparison
    print(f"Created: {args.output}")


def ensure_sic_mapping(args: argparse.Namespace) -> bool:
    """Rebuild the SIC mapping (~1.5GB SEC download) only when it can change the result:
    French changed the FF48 definitions, the mapping file is missing, or a Portfolio/Test
    ticker is not in it (new listing / ticker change). Returns True if it was rebuilt
    (a rebuild also refreshes the factor files)."""
    workdir = Path(args.workdir)
    mapping = workdir / "us_listed_sic_ff48.csv"
    state_path = workdir / MAPPING_STATE_FILE
    state = json.loads(state_path.read_text()) if state_path.exists() else {}

    reasons = []
    definition_date = remote_last_modified(SICCODES_ZIP)
    if definition_date and state.setdefault("siccodes48_last_modified", definition_date) != definition_date:
        download_siccodes(Path(args.ff48_file))
        reasons.append(f"FF48 SIC definitions changed ({definition_date})")
    if not mapping.exists():
        reasons.append("no mapping file")
    else:
        missing = tickers_missing_from_mapping(Path(args.input), mapping, set(state.get("unmappable_tickers", [])))
        if missing:
            reasons.append(f"not in mapping: {', '.join(missing)}")

    rebuilt = False
    if reasons and args.email:
        print(f"Rebuilding SIC mapping ({'; '.join(reasons)})", flush=True)
        quarterly_refresh(build_parser().parse_args([
            "refresh", "--email", args.email, "--ff48-file", args.ff48_file, "--workdir", args.workdir,
        ]))
        still_missing = tickers_missing_from_mapping(Path(args.input), mapping, set())
        if still_missing:
            print(f"Not in the SEC exchange listing even after rebuild (left unmapped, won't retrigger): {', '.join(still_missing)}")
        state.update({
            "siccodes48_last_modified": definition_date or state.get("siccodes48_last_modified"),
            "unmappable_tickers": still_missing,
            "mapping_rebuilt": datetime.now().isoformat(timespec="seconds"),
        })
        rebuilt = True
    elif reasons:
        print(f"SIC mapping rebuild needed ({'; '.join(reasons)}) but no --email configured; skipped")
    else:
        print("SIC mapping up to date; no SEC download needed")
    state_path.write_text(json.dumps(state, indent=2, ensure_ascii=False))
    return rebuilt


def build_report(args: argparse.Namespace) -> Path:
    return create_portfolio_report(build_parser().parse_args([
        "portfolio-report", "--input", args.input,
        "--mapping", str(Path(args.workdir) / "us_listed_sic_ff48.csv"), "--years", str(args.years),
    ]))


def monthly_run(args: argparse.Namespace) -> None:
    """Scheduled on the 1st: SIC check, refresh the French factor files, rebuild the report.
    French publishes a month's data 3-4 weeks after month end, so the run on the 1st
    normally has factor data through the end of the month before last."""
    print(f"===== monthly-run {datetime.now():%Y-%m-%d %H:%M:%S} =====", flush=True)
    if not ensure_sic_mapping(args):
        refresh_factor_files(Path(args.workdir))
    output = build_report(args)
    notify("Portfolio monthly report", f"Report created: {output.name}")


def file_sha256(path: Path) -> str:
    import hashlib
    return hashlib.sha256(path.read_bytes()).hexdigest()


def input_run(args: argparse.Namespace) -> None:
    """Triggered by launchd whenever portfolio_input.xlsx is written. Excel and iCloud can
    touch the file several times per save, so the report is rebuilt only when the file's
    content hash differs from the last processed one. Uses the factor files already on
    disk (no French download); a newly added ticker can still trigger the SIC rebuild."""
    inp = Path(args.input)
    state_path = Path(args.workdir) / INPUT_STATE_FILE
    state = json.loads(state_path.read_text()) if state_path.exists() else {}
    time.sleep(INPUT_SETTLE_SECONDS)  # let Excel finish writing the file
    if not inp.exists():
        return
    digest = file_sha256(inp)
    if digest == state.get("sha256"):
        return  # same content as last processed (repeat event, or only opened/synced)
    print(f"===== input-run {datetime.now():%Y-%m-%d %H:%M:%S} ({inp.name} changed) =====", flush=True)
    ensure_sic_mapping(args)
    output = build_report(args)
    state.update({"sha256": digest, "last_run": datetime.now().isoformat(timespec="seconds")})
    state_path.write_text(json.dumps(state, indent=2))
    notify("Portfolio report", f"{inp.name} changed → report updated: {output.name}")


def remote_last_modified(url: str) -> str | None:
    """Last-Modified header of a French Library file (a HEAD request, no download)."""
    request = urllib.request.Request(url, method="HEAD", headers={"User-Agent": "quarterly-research-automation/1.0"})
    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            return response.headers.get("Last-Modified")
    except Exception as exc:
        print(f"Could not check {url}: {exc}")
        return None


def download_siccodes(destination: Path) -> None:
    request = urllib.request.Request(SICCODES_ZIP, headers={"User-Agent": "quarterly-research-automation/1.0"})
    with urllib.request.urlopen(request, timeout=120) as response, zipfile.ZipFile(io.BytesIO(response.read())) as zf:
        name = next(n for n in zf.namelist() if n.lower().endswith(".txt"))
        tmp = destination.with_suffix(".txt.part")
        tmp.write_bytes(zf.read(name))
    tmp.replace(destination)
    print(f"Overwritten {destination} with the new FF48 SIC definitions")


def tickers_missing_from_mapping(input_path: Path, mapping_path: Path, ignore: set[str]) -> list[str]:
    sheets = pd.read_excel(input_path, sheet_name=None)
    tickers = set()
    for name in ("Portfolio", "Test"):
        if name in sheets:
            tickers |= set(resolve_weights(sheets[name])[0]["Ticker"])
    mapped = set(pd.read_csv(mapping_path, usecols=["ticker"], dtype=str)["ticker"].str.upper().str.strip())
    return sorted(tickers - mapped - ignore)


def notify(title: str, message: str) -> None:
    """macOS notification banner; silently skipped where osascript is unavailable."""
    def quote(text: str) -> str:
        return '"' + text.replace("\\", "\\\\").replace('"', "'") + '"'
    try:
        subprocess.run(["osascript", "-e", f"display notification {quote(message[:200])} with title {quote(title)}"],
                       capture_output=True, timeout=10)
    except (OSError, subprocess.SubprocessError):
        pass


def launch_agent_path(label: str) -> Path:
    return Path.home() / "Library" / "LaunchAgents" / f"{label}.plist"


def write_launch_agent(label: str, command: list[str], log_name: str, calendar=None, watch_paths=None) -> None:
    """(Re)register one launchd job: on a calendar (a run missed while the Mac sleeps fires on
    wake; one missed while it is shut down is skipped) and/or when a watched file changes."""
    project = Path(__file__).resolve().parent
    log_dir = project / "logs"
    log_dir.mkdir(exist_ok=True)
    venv_python = project / ".venv" / "bin" / "python"
    plist = {
        "Label": label,
        "ProgramArguments": [str(venv_python if venv_python.exists() else Path(sys.executable)),
                             str(Path(__file__).resolve()), *command],
        "WorkingDirectory": str(project),
        "StandardOutPath": str(log_dir / log_name),
        "StandardErrorPath": str(log_dir / log_name),
    }
    if calendar is not None:
        plist["StartCalendarInterval"] = calendar
    if watch_paths:
        plist["WatchPaths"] = [str(p) for p in watch_paths]
        plist["ThrottleInterval"] = 15  # at most one start per 15s while Excel saves
    path = launch_agent_path(label)
    path.parent.mkdir(parents=True, exist_ok=True)
    domain = f"gui/{os.getuid()}"
    subprocess.run(["launchctl", "bootout", domain, str(path)], capture_output=True)  # replace any old copy
    with open(path, "wb") as f:
        plistlib.dump(plist, f)
    subprocess.run(["launchctl", "bootstrap", domain, str(path)], check=True)
    print(f"Installed {path}; log -> {log_dir / log_name}")


def install_schedule(args: argparse.Namespace) -> None:
    """Two triggers: monthly-run on the given day of every month, and input-run whenever
    portfolio_input.xlsx is saved."""
    project = Path(__file__).resolve().parent
    email = ["--email", args.email] if args.email else []
    write_launch_agent(LAUNCH_LABEL, ["monthly-run", *email], "monthly.log",
                       calendar={"Day": args.day, "Hour": args.hour, "Minute": args.minute})
    print(f"  monthly-run: day {args.day} of every month at {args.hour:02d}:{args.minute:02d}")
    input_path = project / "portfolio_input.xlsx"
    write_launch_agent(INPUT_WATCH_LABEL, ["input-run", *email], "input.log", watch_paths=[input_path])
    print(f"  input-run: whenever {input_path.name} is saved")
    # The report already reflects the current input, so record it as processed.
    state_path = project / "data" / INPUT_STATE_FILE
    state_path.write_text(json.dumps({"sha256": file_sha256(input_path),
                                      "last_run": datetime.now().isoformat(timespec="seconds")}, indent=2))
    if not args.email:
        print("  (no --email given, so the SIC mapping is never rebuilt automatically)")
    old = launch_agent_path(QUARTERLY_LABEL)  # the former fixed quarterly refresh is now handled by monthly-run
    if old.exists():
        subprocess.run(["launchctl", "bootout", f"gui/{os.getuid()}", str(old)], capture_output=True)
        old.unlink()
        print(f"Removed old quarterly job {old}")


def uninstall_schedule(args: argparse.Namespace) -> None:
    for label in (LAUNCH_LABEL, INPUT_WATCH_LABEL, QUARTERLY_LABEL):
        path = launch_agent_path(label)
        subprocess.run(["launchctl", "bootout", f"gui/{os.getuid()}", str(path)], capture_output=True)
        if path.exists():
            path.unlink()
            print(f"Removed {path}")


def main() -> None:
    args = build_parser().parse_args()
    commands = {
        "refresh": quarterly_refresh,
        "portfolio-report": create_portfolio_report,
        "ticker-export": create_ticker_export,
        "make-template": make_template,
        "factor-refresh": factor_refresh,
        "monthly-run": monthly_run,
        "input-run": input_run,
        "install-schedule": install_schedule,
        "uninstall-schedule": uninstall_schedule,
    }
    try:
        commands[args.command](args)
    except Exception as exc:
        if args.command in ("monthly-run", "input-run", "refresh"):  # the scheduled jobs: surface failures
            notify(f"Portfolio {args.command} FAILED", f"{type(exc).__name__}: {exc}")
        raise
    if args.command == "refresh":
        notify("Portfolio quarterly refresh", "SIC mapping and factor files updated")


if __name__ == "__main__":
    main()
