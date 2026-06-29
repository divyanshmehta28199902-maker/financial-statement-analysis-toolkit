"""
Project 4: Financial Statement Analysis & Valuation Toolkit
===========================================================
Python Module — Core financial analysis engine
Covers: Ratio Analysis | DCF Valuation | Comparable Company Analysis |
        DuPont Analysis | Altman Z-Score | Graham Number

Usage:
    from scripts.financial_toolkit import FinancialAnalyzer
    analyzer = FinancialAnalyzer("Infosys")
    analyzer.run_full_analysis()
"""

import os
import json
import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
from matplotlib.gridspec import GridSpec
from dataclasses import dataclass, field
from typing import Optional

warnings.filterwarnings("ignore")

BASE    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(BASE, "outputs")
os.makedirs(OUT_DIR, exist_ok=True)

# ══════════════════════════════════════════════════════════════
# DATA CLASSES
# ══════════════════════════════════════════════════════════════

@dataclass
class IncomeStatement:
    """Annual income statement in INR Crores"""
    year:               int
    revenue:            float
    cogs:               float
    gross_profit:       float
    ebitda:             float
    depreciation:       float
    ebit:               float
    interest_expense:   float
    ebt:                float
    tax:                float
    net_income:         float
    shares_outstanding: float   # in Crores
    dps:                float = 0.0  # Dividend per share ₹

@dataclass
class BalanceSheet:
    """Annual balance sheet in INR Crores"""
    year:               int
    cash:               float
    current_assets:     float
    total_assets:       float
    current_liabilities: float
    total_debt:         float
    total_liabilities:  float
    shareholders_equity: float
    goodwill:           float = 0.0
    net_fixed_assets:   float = 0.0

@dataclass
class CashFlowStatement:
    """Annual cash flow in INR Crores"""
    year:           int
    cfo:            float   # Operating
    capex:          float   # Capital expenditure (negative)
    fcf:            float   # Free Cash Flow = CFO + Capex
    cff:            float   # Financing
    net_change:     float


# ══════════════════════════════════════════════════════════════
# SAMPLE COMPANY DATA  (Infosys-like IT company)
# ══════════════════════════════════════════════════════════════

def get_sample_company(name: str = "TechCorp India Ltd") -> dict:
    """Returns 5-year financial data mimicking a large-cap IT company."""
    years = [2020, 2021, 2022, 2023, 2024]

    income = [
        IncomeStatement(2020, 90000, 54000, 36000, 22500, 2200, 20300, 500, 19800, 4950, 14850, 420, 34),
        IncomeStatement(2021, 100000, 59000, 41000, 26000, 2400, 23600, 450, 23150, 5788, 17363, 420, 38),
        IncomeStatement(2022, 121000, 71000, 50000, 31500, 2700, 28800, 400, 28400, 7100, 21300, 415, 44),
        IncomeStatement(2023, 146000, 85000, 61000, 38000, 3000, 35000, 380, 34620, 8655, 25965, 415, 50),
        IncomeStatement(2024, 161000, 93000, 68000, 42500, 3300, 39200, 350, 38850, 9713, 29138, 414, 58),
    ]

    balance = [
        BalanceSheet(2020, 8000,  42000, 115000, 18000, 5000,  55000, 60000, 3000, 28000),
        BalanceSheet(2021, 11000, 50000, 130000, 20000, 4500,  57000, 73000, 3000, 27000),
        BalanceSheet(2022, 14000, 62000, 148000, 23000, 4000,  59000, 89000, 3000, 29000),
        BalanceSheet(2023, 16000, 70000, 165000, 26000, 3500,  62000, 103000, 3000, 31000),
        BalanceSheet(2024, 19000, 78000, 182000, 28000, 3000,  64000, 118000, 3000, 33000),
    ]

    cashflow = [
        CashFlowStatement(2020, 18000, -4000, 14000,  -8000, 6000),
        CashFlowStatement(2021, 21000, -4500, 16500,  -9000, 7500),
        CashFlowStatement(2022, 25000, -5500, 19500, -10000, 9500),
        CashFlowStatement(2023, 30000, -6500, 23500, -12000, 11500),
        CashFlowStatement(2024, 34000, -7000, 27000, -14000, 13000),
    ]

    stock_data = {
        "current_price": 1650,   # ₹ per share
        "shares_outstanding": 414,  # Crores
        "beta": 0.85,
        "52w_high": 1950,
        "52w_low": 1300,
    }

    return {"name": name, "years": years,
            "income": income, "balance": balance,
            "cashflow": cashflow, "stock": stock_data}


# ══════════════════════════════════════════════════════════════
# RATIO CALCULATOR
# ══════════════════════════════════════════════════════════════

class RatioAnalyzer:
    """Compute all standard financial ratios from statements."""

    def __init__(self, income: list, balance: list, cashflow: list):
        self.inc = {i.year: i for i in income}
        self.bal = {b.year: b for b in balance}
        self.cf  = {c.year: c for c in cashflow}
        self.years = sorted(self.inc.keys())

    def compute(self) -> pd.DataFrame:
        rows = []
        for y in self.years:
            i, b, c = self.inc[y], self.bal[y], self.cf[y]
            eps = i.net_income / i.shares_outstanding
            bvps = b.shareholders_equity / i.shares_outstanding

            rows.append({
                "Year": y,
                # Profitability
                "Gross Margin %":   round(i.gross_profit / i.revenue * 100, 2),
                "EBITDA Margin %":  round(i.ebitda / i.revenue * 100, 2),
                "EBIT Margin %":    round(i.ebit / i.revenue * 100, 2),
                "Net Margin %":     round(i.net_income / i.revenue * 100, 2),
                "RoA %":            round(i.net_income / b.total_assets * 100, 2),
                "RoE %":            round(i.net_income / b.shareholders_equity * 100, 2),
                "RoCE %":           round(i.ebit / (b.total_assets - b.current_liabilities) * 100, 2),
                # Liquidity
                "Current Ratio":    round(b.current_assets / b.current_liabilities, 2),
                "Cash Ratio":       round(b.cash / b.current_liabilities, 2),
                # Leverage
                "Debt/Equity":      round(i.total_debt / b.shareholders_equity if hasattr(i,'total_debt') else b.total_debt / b.shareholders_equity, 2),
                "Interest Coverage":round(i.ebit / i.interest_expense if i.interest_expense else 0, 2),
                "Net Debt/EBITDA":  round((b.total_debt - b.cash) / i.ebitda, 2),
                # Efficiency
                "Asset Turnover":   round(i.revenue / b.total_assets, 2),
                "FCF Margin %":     round(c.fcf / i.revenue * 100, 2),
                # Per Share
                "EPS ₹":            round(eps, 2),
                "BVPS ₹":           round(bvps, 2),
                "DPS ₹":            i.dps,
                "Payout Ratio %":   round(i.dps / eps * 100, 2) if eps > 0 else 0,
            })

        return pd.DataFrame(rows).set_index("Year")


# ══════════════════════════════════════════════════════════════
# DCF VALUATION
# ══════════════════════════════════════════════════════════════

class DCFValuation:
    """Discounted Cash Flow model with scenario analysis."""

    def __init__(self, base_fcf: float, wacc: float = 0.12,
                 growth_rates: tuple = (0.18, 0.14, 0.10),   # Base case
                 terminal_growth: float = 0.05,
                 shares: float = 414, net_debt: float = -16000):
        """
        base_fcf        : Latest year FCF in ₹ Crores
        wacc            : Weighted Average Cost of Capital
        growth_rates    : (Years 1-2, Years 3-5, Years 6-10) growth
        terminal_growth : Long-run terminal growth rate
        shares          : Shares outstanding in Crores
        net_debt        : Net Debt (negative = net cash)
        """
        self.base_fcf = base_fcf
        self.wacc = wacc
        self.g1, self.g2, self.g3 = growth_rates
        self.tg = terminal_growth
        self.shares = shares
        self.net_debt = net_debt

    def project_fcf(self):
        fcfs, fcf = [], self.base_fcf
        for yr in range(1, 11):
            g = self.g1 if yr <= 2 else (self.g2 if yr <= 5 else self.g3)
            fcf = fcf * (1 + g)
            fcfs.append({"Year": f"Y+{yr}", "FCF_Cr": round(fcf, 0), "Growth": g})
        return pd.DataFrame(fcfs)

    def intrinsic_value(self) -> dict:
        proj = self.project_fcf()
        pv_fcfs = sum(row["FCF_Cr"] / (1 + self.wacc) ** (i+1)
                      for i, row in proj.iterrows())
        terminal_fcf = proj.iloc[-1]["FCF_Cr"] * (1 + self.tg)
        terminal_val = terminal_fcf / (self.wacc - self.tg)
        pv_terminal  = terminal_val / (1 + self.wacc) ** 10
        enterprise_val = pv_fcfs + pv_terminal
        equity_val   = enterprise_val - self.net_debt
        intrinsic_ps = equity_val / self.shares

        return {
            "PV of FCFs (₹Cr)":       round(pv_fcfs, 0),
            "Terminal Value (₹Cr)":    round(terminal_val, 0),
            "PV Terminal Value (₹Cr)": round(pv_terminal, 0),
            "Enterprise Value (₹Cr)":  round(enterprise_val, 0),
            "Equity Value (₹Cr)":      round(equity_val, 0),
            "Intrinsic Value/Share ₹": round(intrinsic_ps, 2),
        }

    def scenario_analysis(self) -> pd.DataFrame:
        scenarios = {
            "Bear":  dict(growth_rates=(0.10, 0.08, 0.06), wacc=0.14, terminal_growth=0.04),
            "Base":  dict(growth_rates=(0.18, 0.14, 0.10), wacc=0.12, terminal_growth=0.05),
            "Bull":  dict(growth_rates=(0.25, 0.20, 0.15), wacc=0.11, terminal_growth=0.06),
        }
        rows = []
        for name, params in scenarios.items():
            dcf = DCFValuation(self.base_fcf, params["wacc"],
                               params["growth_rates"], params["terminal_growth"],
                               self.shares, self.net_debt)
            val = dcf.intrinsic_value()
            rows.append({"Scenario": name, **val})
        return pd.DataFrame(rows)


# ══════════════════════════════════════════════════════════════
# DUPONT ANALYSIS
# ══════════════════════════════════════════════════════════════

def dupont_analysis(income: list, balance: list) -> pd.DataFrame:
    """5-factor DuPont decomposition of RoE."""
    rows = []
    for i, b in zip(income, balance):
        tax_burden    = i.net_income / i.ebt if i.ebt else 0
        interest_burden = i.ebt / i.ebit if i.ebit else 0
        ebit_margin   = i.ebit / i.revenue
        asset_turnover= i.revenue / b.total_assets
        leverage      = b.total_assets / b.shareholders_equity
        roe           = tax_burden * interest_burden * ebit_margin * asset_turnover * leverage

        rows.append({
            "Year":             i.year,
            "Tax Burden":       round(tax_burden, 3),
            "Interest Burden":  round(interest_burden, 3),
            "EBIT Margin":      round(ebit_margin, 3),
            "Asset Turnover":   round(asset_turnover, 3),
            "Leverage":         round(leverage, 3),
            "RoE (DuPont)":     round(roe * 100, 2),
            "RoE (Direct)":     round(i.net_income / b.shareholders_equity * 100, 2),
        })
    return pd.DataFrame(rows).set_index("Year")


# ══════════════════════════════════════════════════════════════
# ALTMAN Z-SCORE
# ══════════════════════════════════════════════════════════════

def altman_zscore(income: list, balance: list) -> pd.DataFrame:
    """Altman Z-Score for non-financial companies (original 1968 model)."""
    rows = []
    for i, b in zip(income, balance):
        mkt_cap = i.shares_outstanding * 1650  # proxy market cap
        working_capital = b.current_assets - b.current_liabilities
        retained_earnings = b.shareholders_equity - b.cash  # simplified proxy

        x1 = working_capital / b.total_assets
        x2 = retained_earnings / b.total_assets
        x3 = i.ebit / b.total_assets
        x4 = mkt_cap / b.total_liabilities
        x5 = i.revenue / b.total_assets

        z  = 1.2*x1 + 1.4*x2 + 3.3*x3 + 0.6*x4 + 1.0*x5
        status = "Safe Zone" if z > 2.99 else ("Grey Zone" if z > 1.81 else "Distress")

        rows.append({"Year": i.year, "X1": round(x1,3), "X2": round(x2,3),
                     "X3": round(x3,3), "X4": round(x4,3), "X5": round(x5,3),
                     "Z-Score": round(z,2), "Status": status})
    return pd.DataFrame(rows).set_index("Year")


# ══════════════════════════════════════════════════════════════
# GRAHAM NUMBER
# ══════════════════════════════════════════════════════════════

def graham_number(eps: float, bvps: float) -> float:
    """Benjamin Graham's intrinsic value formula: √(22.5 × EPS × BVPS)"""
    return round(np.sqrt(22.5 * max(eps, 0) * max(bvps, 0)), 2)


# ══════════════════════════════════════════════════════════════
# MAIN ANALYZER
# ══════════════════════════════════════════════════════════════

class FinancialAnalyzer:

    def __init__(self, company_name: str = "TechCorp India Ltd"):
        data = get_sample_company(company_name)
        self.name    = data["name"]
        self.income  = data["income"]
        self.balance = data["balance"]
        self.cf      = data["cashflow"]
        self.stock   = data["stock"]

    def run_full_analysis(self):
        print(f"\n{'='*60}")
        print(f"  📊 FINANCIAL ANALYSIS: {self.name}")
        print(f"{'='*60}")

        # 1. Ratios
        ratios = RatioAnalyzer(self.income, self.balance, self.cf).compute()
        print("\n📋 KEY FINANCIAL RATIOS")
        print(ratios[["Gross Margin %","EBITDA Margin %","Net Margin %","RoE %","Current Ratio","EPS ₹"]].to_string())

        # 2. DuPont
        dp = dupont_analysis(self.income, self.balance)
        print("\n🔍 DUPONT ANALYSIS")
        print(dp.to_string())

        # 3. Altman Z
        zs = altman_zscore(self.income, self.balance)
        print("\n⚠️  ALTMAN Z-SCORE")
        print(zs[["Z-Score","Status"]].to_string())

        # 4. Graham Number
        latest = self.income[-1]
        eps  = latest.net_income / latest.shares_outstanding
        bvps = self.balance[-1].shareholders_equity / latest.shares_outstanding
        gn   = graham_number(eps, bvps)
        print(f"\n📐 GRAHAM NUMBER: ₹{gn} (Current Price: ₹{self.stock['current_price']})")
        print(f"   → {'UNDERVALUED' if self.stock['current_price'] < gn else 'OVERVALUED'} by Graham metric")

        # 5. DCF
        dcf = DCFValuation(
            base_fcf=self.cf[-1].fcf,
            shares=latest.shares_outstanding,
            net_debt=self.balance[-1].total_debt - self.balance[-1].cash,
        )
        scenarios = dcf.scenario_analysis()
        print("\n💰 DCF SCENARIO ANALYSIS")
        print(scenarios[["Scenario","Intrinsic Value/Share ₹"]].to_string(index=False))
        print(f"   Current Market Price: ₹{self.stock['current_price']}")

        # 6. Charts
        self._plot_dashboard(ratios, scenarios)

        # 7. Export
        self._export(ratios, dp, zs, scenarios)

    def _plot_dashboard(self, ratios: pd.DataFrame, scenarios: pd.DataFrame):
        fig = plt.figure(figsize=(20, 12))
        fig.suptitle(f"{self.name} — Financial Analysis Dashboard", fontsize=16, fontweight="bold")
        gs = GridSpec(2, 4, figure=fig, hspace=0.45, wspace=0.4)

        # Revenue & Net Income
        ax = fig.add_subplot(gs[0, :2])
        yrs = [i.year for i in self.income]
        rev = [i.revenue for i in self.income]
        ni  = [i.net_income for i in self.income]
        x = np.arange(len(yrs)); w = 0.35
        ax.bar(x - w/2, rev, w, label="Revenue", color="#1565C0", alpha=0.8)
        ax.bar(x + w/2, ni,  w, label="Net Income", color="#2E7D32", alpha=0.8)
        ax.set_xticks(x); ax.set_xticklabels(yrs)
        ax.set_title("Revenue vs Net Income (₹Cr)", fontweight="bold")
        ax.yaxis.set_major_formatter(mtick.FuncFormatter(lambda x,_: f"₹{x/1000:.0f}K"))
        ax.legend()

        # Margins trend
        ax2 = fig.add_subplot(gs[0, 2:])
        for col, color in [("Gross Margin %","#1565C0"),("EBITDA Margin %","#E65100"),("Net Margin %","#2E7D32")]:
            ax2.plot(ratios.index, ratios[col], marker="o", label=col, color=color, linewidth=2)
        ax2.set_title("Margin Trends (%)", fontweight="bold"); ax2.legend(fontsize=8)
        ax2.yaxis.set_major_formatter(mtick.PercentFormatter())

        # RoE / RoA
        ax3 = fig.add_subplot(gs[1, 0])
        ax3.plot(ratios.index, ratios["RoE %"], "o-", color="#1565C0", label="RoE", linewidth=2)
        ax3.plot(ratios.index, ratios["RoA %"], "s--", color="#E65100", label="RoA", linewidth=2)
        ax3.set_title("RoE vs RoA (%)", fontweight="bold"); ax3.legend()

        # FCF
        ax4 = fig.add_subplot(gs[1, 1])
        fcf = [c.fcf for c in self.cf]
        ax4.bar(yrs, fcf, color="#2E7D32", alpha=0.8)
        ax4.set_title("Free Cash Flow (₹Cr)", fontweight="bold")
        ax4.yaxis.set_major_formatter(mtick.FuncFormatter(lambda x,_: f"₹{x/1000:.0f}K"))

        # Debt/Equity
        ax5 = fig.add_subplot(gs[1, 2])
        de = ratios["Debt/Equity"]
        ax5.bar(de.index, de.values, color="#6A1B9A", alpha=0.8)
        ax5.set_title("Debt/Equity Ratio", fontweight="bold")

        # DCF Scenario
        ax6 = fig.add_subplot(gs[1, 3])
        colors = {"Bear":"#C62828","Base":"#1565C0","Bull":"#2E7D32"}
        for _, row in scenarios.iterrows():
            ax6.bar(row["Scenario"], row["Intrinsic Value/Share ₹"],
                    color=colors[row["Scenario"]], alpha=0.85)
        ax6.axhline(y=self.stock["current_price"], color="black",
                    linestyle="--", linewidth=1.5, label=f"CMP ₹{self.stock['current_price']}")
        ax6.set_title("DCF Intrinsic Value Scenarios ₹/Share", fontweight="bold")
        ax6.legend(fontsize=8)

        out = os.path.join(OUT_DIR, "financial_dashboard.png")
        plt.savefig(out, dpi=150, bbox_inches="tight")
        plt.close()
        print(f"\n  📊 Dashboard saved: {out}")

    def _export(self, ratios, dupont, zscore, scenarios):
        path = os.path.join(OUT_DIR, f"{self.name.replace(' ','_')}_analysis.xlsx")
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            ratios.to_excel(writer, sheet_name="Key Ratios")
            dupont.to_excel(writer, sheet_name="DuPont Analysis")
            zscore.to_excel(writer, sheet_name="Altman Z-Score")
            scenarios.to_excel(writer, sheet_name="DCF Scenarios", index=False)
            # Raw income
            pd.DataFrame([vars(i) for i in self.income]).to_excel(writer, sheet_name="Income Statement", index=False)
            pd.DataFrame([vars(b) for b in self.balance]).to_excel(writer, sheet_name="Balance Sheet", index=False)
            pd.DataFrame([vars(c) for c in self.cf]).to_excel(writer, sheet_name="Cash Flow", index=False)
        print(f"  📁 Excel workbook saved: {path}")


# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    analyzer = FinancialAnalyzer("TechCorp India Ltd")
    analyzer.run_full_analysis()
