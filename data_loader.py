"""
data_loader.py
--------------
Single source of truth for all data loading and metric computation.
All values are derived directly from the CSV — no hardcoded business metrics.
"""

import os
import pandas as pd


def load_data(filepath: str = None) -> pd.DataFrame:
    """Load and parse the Superstore CSV. Returns a clean DataFrame."""
    if filepath is None:
        filepath = os.path.join(os.path.dirname(__file__), "Sample - Superstore.csv")

    df = pd.read_csv(
        filepath,
        parse_dates=["Order Date", "Ship Date"],
        encoding="latin-1",
    )

    # Derived time columns
    df["Year"] = df["Order Date"].dt.year
    df["Month"] = df["Order Date"].dt.month
    df["YearMonth"] = df["Order Date"].dt.to_period("M").astype(str)

    return df


def compute_metrics(df: pd.DataFrame) -> dict:
    """
    Compute all dashboard metrics from a DataFrame produced by load_data().
    Returns a dict of scalars and DataFrames — never hardcoded values.
    """
    m = {}

    # ── Scalar KPIs ────────────────────────────────────────────────────────────
    m["total_revenue"] = df["Sales"].sum()
    m["total_profit"] = df["Profit"].sum()
    m["profit_margin_pct"] = df["Profit"].sum() / df["Sales"].sum() * 100
    m["total_orders"] = df["Order ID"].nunique()
    m["total_customers"] = df["Customer ID"].nunique()
    m["avg_order_value"] = df["Sales"].sum() / df["Order ID"].nunique()

    # ── Year range (used to populate filter dropdowns) ─────────────────────────
    m["years_sorted"] = sorted(df["Year"].unique().tolist())

    # ── Revenue & Profit by Year ───────────────────────────────────────────────
    yearly = (
        df.groupby("Year")
        .agg(Revenue=("Sales", "sum"), Profit=("Profit", "sum"))
        .reset_index()
    )
    yearly["ProfitMarginPct"] = yearly["Profit"] / yearly["Revenue"] * 100
    m["yearly"] = yearly

    # YoY Revenue growth %
    yearly_sorted = yearly.sort_values("Year").reset_index(drop=True)
    yearly_sorted["RevenueGrowthPct"] = (
        yearly_sorted["Revenue"].pct_change() * 100
    )
    m["yearly_growth"] = yearly_sorted
    # Latest full-year growth (2016→2017)
    last_two = yearly_sorted.dropna(subset=["RevenueGrowthPct"]).tail(1)
    m["latest_yoy_growth_pct"] = (
        last_two["RevenueGrowthPct"].iloc[0] if len(last_two) else None
    )

    # ── Monthly Revenue & Profit trend ────────────────────────────────────────
    monthly = (
        df.groupby("YearMonth")
        .agg(Revenue=("Sales", "sum"), Profit=("Profit", "sum"))
        .reset_index()
        .sort_values("YearMonth")
    )
    m["monthly"] = monthly

    # ── Category metrics ──────────────────────────────────────────────────────
    cat = (
        df.groupby("Category")
        .agg(Revenue=("Sales", "sum"), Profit=("Profit", "sum"))
        .reset_index()
    )
    cat["ProfitMarginPct"] = cat["Profit"] / cat["Revenue"] * 100
    m["category"] = cat

    # ── Sub-Category metrics ──────────────────────────────────────────────────
    subcat = (
        df.groupby("Sub-Category")
        .agg(Revenue=("Sales", "sum"), Profit=("Profit", "sum"))
        .reset_index()
    )
    subcat["ProfitMarginPct"] = subcat["Profit"] / subcat["Revenue"] * 100
    m["subcategory"] = subcat

    # ── Region metrics ────────────────────────────────────────────────────────
    region = (
        df.groupby("Region")
        .agg(Revenue=("Sales", "sum"), Profit=("Profit", "sum"))
        .reset_index()
    )
    region["ProfitMarginPct"] = region["Profit"] / region["Revenue"] * 100
    m["region"] = region

    # ── Segment metrics ───────────────────────────────────────────────────────
    segment = (
        df.groupby("Segment")
        .agg(
            Revenue=("Sales", "sum"),
            Profit=("Profit", "sum"),
            Orders=("Order ID", "nunique"),
            Customers=("Customer ID", "nunique"),
        )
        .reset_index()
    )
    segment["ProfitMarginPct"] = segment["Profit"] / segment["Revenue"] * 100
    m["segment"] = segment

    # ── Top 10 Products ───────────────────────────────────────────────────────
    prod = (
        df.groupby("Product Name")
        .agg(Revenue=("Sales", "sum"), Profit=("Profit", "sum"))
        .reset_index()
    )
    m["top10_products_by_revenue"] = prod.nlargest(10, "Revenue").reset_index(drop=True)
    m["top10_products_by_profit"] = prod.nlargest(10, "Profit").reset_index(drop=True)

    # ── Top 10 Customers ──────────────────────────────────────────────────────
    cust = (
        df.groupby(["Customer ID", "Customer Name"])
        .agg(Revenue=("Sales", "sum"), Profit=("Profit", "sum"))
        .reset_index()
    )
    m["top10_customers"] = cust.nlargest(10, "Revenue").reset_index(drop=True)

    # ── Row-level scatter data ────────────────────────────────────────────────
    # Keep only the columns needed for scatter charts to reduce memory
    m["scatter_data"] = df[["Sales", "Profit", "Discount", "Category", "Sub-Category"]].copy()

    return m
