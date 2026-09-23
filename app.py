"""
app.py
------
Plotly Dash multi-page dashboard entry point.
Run:  python app.py   →   http://127.0.0.1:8050
"""

import dash
import dash_bootstrap_components as dbc
from dash import dcc, html

# Pre-load data once at startup so all pages share the same objects
import data_loader as dl

DF = dl.load_data()
METRICS = dl.compute_metrics(DF)

app = dash.Dash(
    __name__,
    use_pages=True,
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    suppress_callback_exceptions=True,
)
app.title = "Retail Sales BI Dashboard"

# ── Navigation bar ─────────────────────────────────────────────────────────────
navbar = dbc.Navbar(
    dbc.Container(
        [
            dbc.NavbarBrand(
                "🛒 Retail Sales BI Dashboard",
                className="fw-bold fs-5",
                style={"color": "#ffffff"},
            ),
            dbc.Nav(
                [
                    dbc.NavItem(
                        dbc.NavLink(
                            "Executive Overview",
                            href="/",
                            active="exact",
                            style={"color": "#e0e0e0"},
                        )
                    ),
                    dbc.NavItem(
                        dbc.NavLink(
                            "Sales & Products",
                            href="/sales-product",
                            active="exact",
                            style={"color": "#e0e0e0"},
                        )
                    ),
                    dbc.NavItem(
                        dbc.NavLink(
                            "Customer & Risk",
                            href="/customer-risk",
                            active="exact",
                            style={"color": "#e0e0e0"},
                        )
                    ),
                ],
                navbar=True,
                className="ms-auto",
            ),
        ],
        fluid=True,
    ),
    color="#1f3f6e",
    dark=True,
    className="mb-3",
)

# ── Sub-header strip ───────────────────────────────────────────────────────────
subheader = dbc.Container(
    dbc.Row(
        dbc.Col(
            html.P(
                "IBM SkillsBuild Data Analytics with AI Internship — Final Project  |  "
                "Dataset: Sample - Superstore (2014–2017)  |  9,994 transactions",
                className="text-muted small mb-2",
            )
        )
    ),
    fluid=True,
)

# ── Root layout ────────────────────────────────────────────────────────────────
app.layout = html.Div(
    [
        dcc.Location(id="url"),
        navbar,
        subheader,
        dash.page_container,
    ]
)

if __name__ == "__main__":
    app.run(debug=False, host="127.0.0.1", port=8050)
