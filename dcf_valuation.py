# ============================================================
#  DCF VALUATION MODEL
#  Author: [Your Name]
#  Description: Builds a Discounted Cash Flow (DCF) model
#               for any public company using Yahoo Finance data.
#               Includes sensitivity analysis and Excel export.
# ============================================================

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import warnings
warnings.filterwarnings("ignore")

try:
    import yfinance as yf
    YF_OK = True
except ImportError:
    YF_OK = False
    print("Install yfinance: pip install yfinance")

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False
    print("Install openpyxl: pip install openpyxl")


# ============================================================
# CONFIGURATION — change the ticker to value any company!
# ============================================================

TICKER           = "AAPL"       # Stock ticker (e.g. MSFT, TSLA, JPM)
PROJECTION_YEARS = 5            # Years of free cash flow projection
WACC             = 0.10         # Weighted Average Cost of Capital (10%)
TERMINAL_GROWTH  = 0.03         # Long-run growth rate (3% = GDP-like)
REVENUE_GROWTH   = 0.08         # Assumed annual revenue growth rate
FCF_MARGIN       = 0.25         # Free Cash Flow as % of revenue


# ============================================================
# STEP 1: Fetch company data from Yahoo Finance
# ============================================================

def fetch_company_data(ticker):
    if not YF_OK:
        print("⚠️  yfinance not available. Using demo data.")
        return None, None

    print(f"\n📥 Fetching data for {ticker}...")
    stock     = yf.Ticker(ticker)
    info      = stock.info
    financials = stock.financials

    name           = info.get("longName", ticker)
    market_cap     = info.get("marketCap", 0)
    shares_out     = info.get("sharesOutstanding", 0)
    current_price  = info.get("currentPrice", 0)
    sector         = info.get("sector", "N/A")

    # Try to get last year's revenue from financials
    try:
        last_revenue = financials.loc["Total Revenue"].iloc[0]
    except Exception:
        last_revenue = info.get("totalRevenue", 300_000_000_000)

    summary = {
        "name":          name,
        "ticker":        ticker,
        "market_cap":    market_cap,
        "shares_out":    shares_out,
        "current_price": current_price,
        "sector":        sector,
        "last_revenue":  last_revenue,
    }

    print(f"✅ {name} ({sector})")
    print(f"   Current price: ${current_price:,.2f}  |  Market cap: ${market_cap/1e9:.1f}B")
    return summary, stock


def demo_data():
    """Fallback demo if yfinance is unavailable."""
    return {
        "name":          "Apple Inc.",
        "ticker":        "AAPL",
        "market_cap":    3_000_000_000_000,
        "shares_out":    15_400_000_000,
        "current_price": 195.00,
        "sector":        "Technology",
        "last_revenue":  383_285_000_000,
    }


# ============================================================
# STEP 2: Project Free Cash Flows
# ============================================================

def project_fcf(last_revenue, years, growth_rate, fcf_margin):
    """
    Projects revenue and Free Cash Flow (FCF) for n years.
    FCF = Revenue × FCF margin (simplified assumption).
    """
    revenues = []
    fcfs     = []

    rev = last_revenue
    for y in range(1, years + 1):
        rev = rev * (1 + growth_rate)
        fcf = rev * fcf_margin
        revenues.append(rev)
        fcfs.append(fcf)

    df = pd.DataFrame({
        "Year":     [f"Year {i}" for i in range(1, years + 1)],
        "Revenue":  revenues,
        "FCF":      fcfs,
    })

    print(f"\n📊 Projected Free Cash Flows (growth: {growth_rate*100:.0f}%/yr):")
    for _, row in df.iterrows():
        print(f"   {row['Year']}: Revenue ${row['Revenue']/1e9:.1f}B  |  FCF ${row['FCF']/1e9:.1f}B")

    return df


# ============================================================
# STEP 3: Calculate intrinsic value
# ============================================================

def calculate_dcf(fcf_df, wacc, terminal_growth, shares_out):
    """
    Discounts projected FCFs back to present value.
    Adds terminal value (Gordon Growth Model).
    Divides by shares outstanding for intrinsic price.
    """
    fcfs = fcf_df["FCF"].values

    # Present value of each FCF
    pv_fcfs = [fcf / (1 + wacc) ** (i + 1) for i, fcf in enumerate(fcfs)]
    sum_pv   = sum(pv_fcfs)

    # Terminal value: last FCF grown at terminal rate, capitalised and discounted
    terminal_value    = fcfs[-1] * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_terminal       = terminal_value / (1 + wacc) ** len(fcfs)

    enterprise_value  = sum_pv + pv_terminal
    intrinsic_value   = enterprise_value / shares_out if shares_out > 0 else 0

    print(f"\n💰 DCF Valuation Results:")
    print(f"   PV of FCFs:        ${sum_pv/1e9:.1f}B")
    print(f"   PV Terminal Value: ${pv_terminal/1e9:.1f}B  ({pv_terminal/enterprise_value*100:.0f}% of EV)")
    print(f"   Enterprise Value:  ${enterprise_value/1e9:.1f}B")
    print(f"   Intrinsic Value:   ${intrinsic_value:.2f} per share")

    return {
        "pv_fcfs":          sum_pv,
        "pv_fcfs_list":     pv_fcfs,
        "pv_terminal":      pv_terminal,
        "enterprise_value": enterprise_value,
        "intrinsic_value":  intrinsic_value,
    }


# ============================================================
# STEP 4: Sensitivity analysis — WACC vs Terminal Growth
# ============================================================

def sensitivity_analysis(fcf_df, base_wacc, base_tg, shares_out):
    """
    Creates a matrix of intrinsic values across
    different WACC and terminal growth rate assumptions.
    """
    waccs    = [base_wacc - 0.02, base_wacc - 0.01, base_wacc,
                base_wacc + 0.01, base_wacc + 0.02]
    tgrowths = [base_tg - 0.01, base_tg, base_tg + 0.01,
                base_tg + 0.02, base_tg + 0.03]
    fcfs     = fcf_df["FCF"].values

    rows = {}
    for tg in tgrowths:
        row = []
        for w in waccs:
            pv    = sum(f / (1 + w) ** (i + 1) for i, f in enumerate(fcfs))
            tv    = fcfs[-1] * (1 + tg) / (w - tg) / (1 + w) ** len(fcfs)
            iv    = (pv + tv) / shares_out if shares_out > 0 else 0
            row.append(round(iv, 2))
        rows[f"TG {tg*100:.0f}%"] = row

    sa_df = pd.DataFrame(rows, index=[f"WACC {w*100:.0f}%" for w in waccs]).T
    print("\n📋 Sensitivity Table (Intrinsic Value $):")
    print(sa_df.to_string())
    return sa_df, waccs, tgrowths


# ============================================================
# STEP 5: Plot results
# ============================================================

def plot_dcf(fcf_df, dcf_results, sensitivity_df, summary, waccs, tgrowths):
    fig, axes = plt.subplots(1, 3, figsize=(16, 6))
    fig.patch.set_facecolor("#0d1117")

    for ax in axes:
        ax.set_facecolor("#0d1117")
        ax.tick_params(colors="white")
        for spine in ax.spines.values():
            spine.set_edgecolor("#333")
        ax.grid(color="#1e1e2e", linewidth=0.5, linestyle="--")

    # Chart 1: FCF waterfall
    years = fcf_df["Year"].tolist() + ["Terminal\nValue"]
    vals  = [v / 1e9 for v in dcf_results["pv_fcfs_list"]] + [dcf_results["pv_terminal"] / 1e9]
    colors = ["#4A90D9"] * len(dcf_results["pv_fcfs_list"]) + ["#F5A623"]
    axes[0].bar(years, vals, color=colors, edgecolor="#0d1117", width=0.6)
    axes[0].set_title("PV of Cash Flows ($B)", color="white", fontsize=11)
    axes[0].set_ylabel("$B", color="white")
    for i, v in enumerate(vals):
        axes[0].text(i, v + 0.5, f"${v:.0f}B", ha="center", color="white", fontsize=9)

    # Chart 2: Enterprise value breakdown (pie)
    ev_parts = [dcf_results["pv_fcfs"], dcf_results["pv_terminal"]]
    labels   = [f"PV FCFs\n${dcf_results['pv_fcfs']/1e9:.0f}B",
                f"Terminal Value\n${dcf_results['pv_terminal']/1e9:.0f}B"]
    axes[1].pie(ev_parts, labels=labels, colors=["#4A90D9", "#F5A623"],
                autopct="%1.0f%%", textprops={"color": "white", "fontsize": 10},
                wedgeprops={"edgecolor": "#0d1117", "linewidth": 2})
    axes[1].set_title("Enterprise Value Breakdown", color="white", fontsize=11)

    # Chart 3: Sensitivity heatmap
    sa_vals = sensitivity_df.values.astype(float)
    im = axes[2].imshow(sa_vals, cmap="RdYlGn", aspect="auto")
    axes[2].set_xticks(range(len(waccs)))
    axes[2].set_xticklabels([f"{w*100:.0f}%" for w in waccs], color="white", fontsize=8)
    axes[2].set_yticks(range(len(tgrowths)))
    axes[2].set_yticklabels([f"{t*100:.0f}%" for t in tgrowths], color="white", fontsize=8)
    axes[2].set_xlabel("WACC", color="white")
    axes[2].set_ylabel("Terminal Growth", color="white")
    axes[2].set_title("Sensitivity: Intrinsic Value ($)", color="white", fontsize=11)
    for i in range(sa_vals.shape[0]):
        for j in range(sa_vals.shape[1]):
            axes[2].text(j, i, f"${sa_vals[i,j]:.0f}", ha="center",
                         va="center", color="white", fontsize=8)

    iv   = dcf_results["intrinsic_value"]
    cp   = summary["current_price"]
    updn = ((iv - cp) / cp * 100) if cp > 0 else 0
    fig.suptitle(
        f"{summary['name']} DCF Valuation  |  Intrinsic Value: ${iv:.2f}  "
        f"({'▲' if updn >= 0 else '▼'} {abs(updn):.1f}% vs ${cp:.2f} market price)",
        color="white", fontsize=13, fontweight="bold", y=1.02
    )

    plt.tight_layout()
    plt.savefig("dcf_valuation.png", dpi=150, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.show()
    print("\n💾 Chart saved as dcf_valuation.png")


# ============================================================
# STEP 6: Export to Excel
# ============================================================

def export_excel(summary, fcf_df, dcf_results, sensitivity_df):
    if not EXCEL_OK:
        print("⚠️  Skipping Excel export.")
        return

    fname = "dcf_valuation_report.xlsx"
    with pd.ExcelWriter(fname, engine="openpyxl") as writer:

        # Summary sheet
        iv  = dcf_results["intrinsic_value"]
        cp  = summary["current_price"]
        updn = ((iv - cp) / cp * 100) if cp > 0 else 0
        summ = pd.DataFrame({
            "Metric": ["Company", "Ticker", "Sector", "Current Price ($)",
                       "Intrinsic Value ($)", "Upside / Downside (%)",
                       "Enterprise Value ($B)", "WACC (%)", "Terminal Growth (%)"],
            "Value":  [summary["name"], summary["ticker"], summary["sector"],
                       round(cp, 2), round(iv, 2), round(updn, 1),
                       round(dcf_results["enterprise_value"]/1e9, 1),
                       round(WACC*100, 1), round(TERMINAL_GROWTH*100, 1)]
        })
        summ.to_excel(writer, sheet_name="DCF Summary", index=False)

        # FCF projections
        fcf_export = fcf_df.copy()
        fcf_export["Revenue ($B)"] = (fcf_export["Revenue"] / 1e9).round(2)
        fcf_export["FCF ($B)"]     = (fcf_export["FCF"] / 1e9).round(2)
        fcf_export["PV of FCF ($B)"] = [round(v/1e9, 2) for v in dcf_results["pv_fcfs_list"]]
        fcf_export[["Year","Revenue ($B)","FCF ($B)","PV of FCF ($B)"]].to_excel(
            writer, sheet_name="FCF Projections", index=False)

        # Sensitivity
        sensitivity_df.to_excel(writer, sheet_name="Sensitivity Analysis")

    print(f"📁 Excel report saved as {fname}")


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 55)
    print("   DCF VALUATION MODEL")
    print(f"   Valuing {TICKER} · WACC {WACC*100:.0f}% · TG {TERMINAL_GROWTH*100:.0f}%")
    print("=" * 55)

    if YF_OK:
        summary, _ = fetch_company_data(TICKER)
        if summary is None:
            summary = demo_data()
    else:
        summary = demo_data()

    fcf_df     = project_fcf(summary["last_revenue"], PROJECTION_YEARS,
                              REVENUE_GROWTH, FCF_MARGIN)
    dcf_results = calculate_dcf(fcf_df, WACC, TERMINAL_GROWTH, summary["shares_out"])
    sa_df, waccs, tgrowths = sensitivity_analysis(fcf_df, WACC, TERMINAL_GROWTH,
                                                   summary["shares_out"])

    iv = dcf_results["intrinsic_value"]
    cp = summary["current_price"]
    if cp > 0:
        verdict = "UNDERVALUED ✅" if iv > cp else "OVERVALUED ⚠️"
        print(f"\n🏆 Verdict: {verdict}  (IV ${iv:.2f} vs Market ${cp:.2f})")

    plot_dcf(fcf_df, dcf_results, sa_df, summary, waccs, tgrowths)
    export_excel(summary, fcf_df, dcf_results, sa_df)

    print("\n✅ All done!")
    print("=" * 55)


if __name__ == "__main__":
    main()
