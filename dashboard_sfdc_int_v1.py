"""
SAP Concur INT SMB Client Sales — Live SFDC Dashboard
Markets: Canada (CAD) · United Kingdom (GBP) · Australia (AUD)

Run:  streamlit run dashboard_sfdc_int_v1.py
"""

import io, json, re, time as _time, warnings
from calendar import monthrange
from datetime import datetime
from pathlib import Path

import requests as _requests

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SAP Concur | INT SMB Client Sales Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# BRAND
# ─────────────────────────────────────────────────────────────────────────────
C = dict(blue="#0070F2", light_blue="#4CB1FF", dark_blue="#00144A",
         green="#188918", red="#BB0000", amber="#E8A000", yellow="#F0AB00",
         grey="#E1E2E6", dark_grey="#6A6D73", bg="#FFFFFF", bg_subtle="#F5F6F7")

MARKET_COLORS = {
    "Canada":         C["blue"],
    "United Kingdom": C["dark_blue"],
    "Australia":      C["light_blue"],
}
BUCKET_ORDER  = ["<50%", "50-75%", "75-100%", "100-120%", "120%+"]
BUCKET_COLORS = {"<50%": C["red"], "50-75%": C["amber"], "75-100%": C["yellow"],
                 "100-120%": "#5DB533", "120%+": C["green"]}
MONTH_NAMES   = ["January","February","March","April","May","June",
                 "July","August","September","October","November","December"]
MONTH_ABBREV  = ["Jan","Feb","Mar","Apr","May","Jun",
                 "Jul","Aug","Sep","Oct","Nov","Dec"]

# ─────────────────────────────────────────────────────────────────────────────
# INT ORG STRUCTURE
# ─────────────────────────────────────────────────────────────────────────────

# Currency per market (used in SFDC reports and quota file)
MARKET_CURRENCY = {
    "Canada":         "CAD",
    "United Kingdom": "GBP",
    "Australia":      "AUD",
}
CURRENCY_SYMBOL = {"CAD": "C$", "GBP": "£", "AUD": "A$", "USD": "$", "EUR": "€", "AED": "AED "}

# 2026 Budget Exchange Rates — local → USD  (source: 2026 Budget Rates.xlsx, "to USD" sheet)
# Used only for the display layer: rep-level monetary amounts are converted from
# USD back to the rep's local currency before rendering.  All internal calculations
# (attainment, quotas, performance totals) remain in USD throughout.
FX_RATES: dict[str, float] = {
    "AUD": 0.6629213,   # 1 AUD = 0.6629213 USD
    "CAD": 0.7284,      # 1 CAD = 0.7284 USD
    "GBP": 1.32584,     # 1 GBP = 1.32584 USD
    "EUR": 1.18,        # 1 EUR ≈ 1.18 USD  (2026 budget rate)
    "AED": 0.2731481,   # 1 AED = 0.2731481 USD
    "USD": 1.0,
}

# Concur Team → FLSM (first-level sales manager / equivalent of RSD)
TEAM_LEADER_MAP = {
    "Canada SMB Client Sales Key":        "Lesley Nunes",
    "Canada SMB Client Sales Premier":    "Lesley Nunes",
    "Canada SMB Client Sales Strategic":  "Lesley Nunes",
    "UK SMB Client Sales Key":            "Katie Brown",
    "UK SMB Client Sales Strategic":      "Bret Edis",
    "UK SMB Client Sales Premier":        "Bret Edis",
    "UK Mid Market":                      "Jessica Brown",
    "UK National":                        "Ryan Headington",
    "UK Premier":                         "Ryan Headington",
    "UK Ireland":                         "Nic Henney",
    "Australia SMB Client Sales":         "Peter Soukos",
    "Australia Mid Market":               "Hamish Tebbutt",
}

# FLSM → Director (one level up)
FLSM_DIRECTOR = {
    "Lesley Nunes":     "Brian Veloso",
    "Katie Brown":      "Adam Bazeley",
    "Bret Edis":        "Adam Bazeley",
    "Jessica Brown":    "Angus Milledge",
    "Ryan Headington":  "Angus Milledge",
    "Nic Henney":       "Angus Milledge",
    "Peter Soukos":     "Fabian Calle",
    "Hamish Tebbutt":   "Fabian Calle",
}

# Director → Market
DIRECTOR_MARKET = {
    "Brian Veloso":   "Canada",
    "Adam Bazeley":   "United Kingdom",
    "Angus Milledge": "United Kingdom",
    "Fabian Calle":   "Australia",
}

# Concur Team → Market
TEAM_MARKET = {
    "Canada SMB Client Sales Key":        "Canada",
    "Canada SMB Client Sales Premier":    "Canada",
    "Canada SMB Client Sales Strategic":  "Canada",
    "UK SMB Client Sales Key":            "United Kingdom",
    "UK SMB Client Sales Strategic":      "United Kingdom",
    "UK SMB Client Sales Premier":        "United Kingdom",
    "UK Mid Market":                      "United Kingdom",
    "UK National":                        "United Kingdom",
    "UK Premier":                         "United Kingdom",
    "UK Ireland":                         "United Kingdom",
    "Australia SMB Client Sales":         "Australia",
    "Australia Mid Market":               "Australia",
}

# Market leaders (used for org-level attainment attribution)
MARKET_LEADER = {
    "Canada":         "Brian Veloso",
    "United Kingdom": "Cassie Petrie",
    "Australia":      "Fabian Calle",
}

# Name normalisations (add as needed when SFDC names differ from quota file)
NAME_MAP = {
    "Conor Tomlinson": "conor tomlinson",   # quota file uses lowercase
}

# ─────────────────────────────────────────────────────────────────────────────
# EMBEDDED QUOTA DATA  (PR monthly amounts in USD, FY 2026)
# Source: 2026 SMB quotas in USD 9.17.26.xlsx — GTM Ops Plan Summary Concur
# PR columns are already pro-rated for mid-year starters (zeros before first quota month).
# The 'currency' field indicates each rep's LOCAL currency for display purposes only.
# All attainment calculations use these USD values alongside USD-converted SFDC fields.
# ─────────────────────────────────────────────────────────────────────────────
QUOTA_DATA = [
    {'name': 'Andrew Cooksley', 'market': 'Australia', 'team': 'Australia SMB Client Sales', 'currency': 'AUD', 'pr': [23869.06, 28643.23, 42964.41, 28231.22, 33877.38, 50815.86, 29551.99, 35462.64, 53193.75, 30427.99, 36513.51, 54770.05]},
    {'name': 'Kate Hulmston', 'market': 'Australia', 'team': 'Australia SMB Client Sales', 'currency': 'AUD', 'pr': [38933.04, 46720.24, 70079.63, 46048.21, 55257.71, 82886.19, 48202.5, 57843.44, 86764.79, 49631.39, 59557.52, 89335.92]},
    {'name': 'Amanda Player', 'market': 'Australia', 'team': 'Australia SMB Client Sales', 'currency': 'AUD', 'pr': [38933.04, 46720.24, 70079.63, 46048.21, 55257.71, 82886.19, 48202.5, 57843.44, 86764.79, 49631.39, 59557.52, 89335.92]},
    {'name': 'Steve Kavanagh', 'market': 'Australia', 'team': 'Australia SMB Client Sales', 'currency': 'AUD', 'pr': [23869.06, 28643.23, 42964.41, 28231.22, 33877.38, 50815.86, 29551.99, 35462.64, 53193.75, 30427.99, 36513.51, 54770.05]},
    {'name': 'Claire van der Vegt', 'market': 'Australia', 'team': 'Australia SMB Client Sales', 'currency': 'AUD', 'pr': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 35462.35, 53193.3, 30427.75, 36513.21, 54769.59]},
    {'name': 'Peter Axon', 'market': 'Australia', 'team': 'Australia Mid Market', 'currency': 'AUD', 'pr': [9775.04, 11730.2, 17595.12, 11561.47, 13873.73, 20810.5, 12102.36, 14522.94, 21784.32, 12461.11, 14953.3, 22429.86]},
    {'name': 'Darcy Penman', 'market': 'Australia', 'team': 'Australia Mid Market', 'currency': 'AUD', 'pr': [9775.04, 11730.2, 17595.12, 11561.47, 13873.73, 20810.5, 12102.36, 14522.94, 21784.32, 12461.11, 14953.3, 22429.86]},
    {'name': 'Nick Bright', 'market': 'United Kingdom', 'team': 'UK Ireland', 'currency': 'GBP', 'pr': [12824.6, 15389.51, 23084.29, 14573.41, 17488.09, 26232.14, 15156.35, 18187.62, 27281.43, 15739.28, 18887.15, 28330.71]},
    {'name': 'Danny Gloyne', 'market': 'United Kingdom', 'team': 'UK Ireland', 'currency': 'GBP', 'pr': [12824.6, 15389.52, 23084.29, 14573.4, 17488.09, 26232.14, 15156.35, 18187.62, 27281.43, 15739.28, 18887.15, 28330.71]},
    {'name': 'Olivia Allen', 'market': 'United Kingdom', 'team': 'UK Ireland', 'currency': 'GBP', 'pr': [12824.6, 15389.52, 23084.29, 14573.41, 17488.09, 26232.14, 15156.35, 18187.62, 27281.42, 15739.28, 18887.15, 28330.71]},
    {'name': 'Bryn Cowling', 'market': 'United Kingdom', 'team': 'UK Mid Market', 'currency': 'GBP', 'pr': [9908.14, 11889.88, 17834.73, 11434.1, 13720.92, 20581.39, 12153.36, 14584.03, 21876.05, 12728.76, 15274.62, 22912.22]},
    {'name': 'Hannah White', 'market': 'United Kingdom', 'team': 'UK Mid Market', 'currency': 'GBP', 'pr': [9908.14, 11889.88, 17834.73, 11434.1, 13720.93, 20581.39, 12153.36, 14584.03, 21876.05, 12728.76, 15274.61, 22912.22]},
    {'name': 'Loreena Maguet', 'market': 'United Kingdom', 'team': 'UK Mid Market', 'currency': 'EUR', 'pr': [9908.14, 11889.88, 17834.73, 11434.1, 13720.93, 20581.38, 12153.36, 14584.03, 21876.05, 12728.76, 15274.62, 22912.22]},
    {'name': 'Calvin Nisban', 'market': 'United Kingdom', 'team': 'UK Mid Market', 'currency': 'GBP', 'pr': [9908.14, 11889.88, 17834.73, 11434.1, 13720.93, 20581.38, 12153.36, 14584.03, 21876.05, 12728.76, 15274.62, 22912.22]},
    {'name': 'Lily Shaw', 'market': 'United Kingdom', 'team': 'UK Mid Market', 'currency': 'GBP', 'pr': [9908.14, 11889.88, 17834.73, 11434.1, 13720.93, 20581.39, 12153.36, 14584.03, 21876.05, 12728.75, 15274.62, 22912.22]},
    {'name': 'Tom Evans', 'market': 'United Kingdom', 'team': 'UK Mid Market', 'currency': 'GBP', 'pr': [9908.14, 11889.86, 17834.73, 11434.1, 13720.93, 20581.39, 12153.36, 14584.03, 21876.05, 12728.76, 15274.62, 22912.22]},
    {'name': 'Lydia Morrell', 'market': 'United Kingdom', 'team': 'UK Mid Market', 'currency': 'GBP', 'pr': [9908.14, 11889.88, 17834.73, 11434.1, 13720.93, 20581.39, 12153.35, 14584.03, 21876.05, 12728.76, 15274.62, 22912.22]},
    {'name': 'Jake Jenkins', 'market': 'United Kingdom', 'team': 'UK Mid Market', 'currency': 'GBP', 'pr': [9908.14, 11889.88, 17834.73, 11434.1, 13720.92, 20581.39, 12153.36, 14584.03, 21876.05, 12728.76, 15274.62, 22912.22]},
    {'name': 'James Hirst', 'market': 'United Kingdom', 'team': 'UK Mid Market', 'currency': 'GBP', 'pr': [0.0, 11889.88, 17834.73, 11434.09, 13720.93, 20581.39, 12153.36, 14584.03, 21876.05, 12728.76, 15274.62, 22912.22]},
    {'name': 'Ben Chandiram', 'market': 'United Kingdom', 'team': 'UK Premier', 'currency': 'GBP', 'pr': [18791.4, 22549.89, 33824.67, 21685.49, 26022.59, 39033.89, 23049.6, 27659.53, 41489.28, 24140.9, 28969.28, 43454.45]},
    {'name': 'Jack Morris', 'market': 'United Kingdom', 'team': 'UK Premier', 'currency': 'GBP', 'pr': [18791.4, 22549.89, 33824.67, 21685.48, 26022.59, 39033.89, 23049.6, 27659.53, 41489.29, 24140.9, 28969.28, 43454.45]},
    {'name': 'Sachin Wilde', 'market': 'United Kingdom', 'team': 'UK Premier', 'currency': 'GBP', 'pr': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 23049.58, 27659.53, 41489.29, 24140.9, 28969.28, 43454.45]},
    {'name': 'Alastair King', 'market': 'United Kingdom', 'team': 'UK National', 'currency': 'GBP', 'pr': [16422.74, 19707.48, 29561.07, 18952.03, 22742.44, 34113.67, 20144.2, 24173.04, 36259.57, 21097.93, 25317.7, 37977.01]},
    {'name': "Glen O'Brien", 'market': 'United Kingdom', 'team': 'UK National', 'currency': 'GBP', 'pr': [16422.75, 19707.48, 29561.07, 18952.02, 22742.44, 34113.67, 20144.2, 24173.04, 36259.57, 21097.93, 25317.7, 37977.01]},
    {'name': 'James Hooker', 'market': 'United Kingdom', 'team': 'UK National', 'currency': 'GBP', 'pr': [16422.75, 19707.48, 29561.06, 18952.03, 22742.44, 34113.67, 20144.2, 24173.04, 36259.57, 21097.93, 25317.7, 37977.01]},
    {'name': 'richard vines', 'market': 'United Kingdom', 'team': 'UK National', 'currency': 'GBP', 'pr': [16422.74, 19707.48, 29561.07, 18952.03, 22742.44, 34113.67, 20144.2, 24173.04, 36259.57, 21097.93, 25317.7, 37977.01]},
    {'name': 'Alan Donohoe', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Premier', 'currency': 'EUR', 'pr': [46994.46, 56393.88, 84590.38, 54232.13, 65078.56, 97617.86, 57643.57, 69172.29, 103758.43, 60372.71, 72447.79, 108673.0]},
    {'name': 'Lydia Holloway', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Premier', 'currency': 'AED', 'pr': [46994.46, 56393.88, 84590.38, 54232.13, 65078.56, 97617.86, 57643.57, 69172.29, 103758.43, 60372.71, 72447.79, 108673.0]},
    {'name': 'Kammie Flitton', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Premier', 'currency': 'GBP', 'pr': [46994.46, 56393.88, 84590.38, 54232.13, 65078.56, 97617.86, 57643.57, 69172.29, 103758.43, 60372.71, 72447.79, 108673.0]},
    {'name': 'Karl Perkins', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Strategic', 'currency': 'GBP', 'pr': [35104.96, 42126.34, 63189.18, 40511.51, 48613.81, 72920.73, 43059.86, 51671.83, 77507.75, 45098.53, 54118.64, 81178.94]},
    {'name': 'Izabella Krawczyk Patel', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Strategic', 'currency': 'GBP', 'pr': [35104.95, 42126.35, 63189.18, 40511.51, 48613.81, 72920.73, 43059.86, 51671.83, 77507.75, 45098.53, 54118.64, 81178.94]},
    {'name': 'Joanna Blackmore', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Strategic', 'currency': 'GBP', 'pr': [35104.95, 42126.34, 63189.18, 40511.51, 48613.81, 72920.73, 43059.86, 51671.83, 77507.75, 45098.53, 54118.64, 81178.96]},
    {'name': 'Michael Benn', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Strategic', 'currency': 'GBP', 'pr': [35104.95, 42126.34, 63189.18, 40511.51, 48613.81, 72920.73, 43059.86, 51671.83, 77507.75, 45098.53, 54118.65, 81178.94]},
    {'name': 'Ryan Hale', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Strategic', 'currency': 'GBP', 'pr': [35104.87, 42126.24, 63189.03, 40511.41, 48613.71, 72920.55, 43059.75, 51671.72, 77507.56, 45098.43, 54118.51, 81178.76]},
    {'name': 'Charlie Mason', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Key', 'currency': 'GBP', 'pr': [26515.11, 31818.45, 47727.42, 30598.75, 36718.5, 55077.76, 32523.54, 39028.25, 58542.37, 34063.37, 40876.35, 61315.27]},
    {'name': 'Adrian Sage', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Key', 'currency': 'GBP', 'pr': [26515.13, 31818.45, 47727.42, 30598.75, 36718.5, 55077.76, 32523.54, 39028.25, 58542.37, 34063.37, 40876.34, 61315.27]},
    {'name': 'Zak Jones', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Key', 'currency': 'GBP', 'pr': [26515.13, 31818.45, 47727.41, 30598.75, 36718.49, 55077.76, 32523.54, 39028.25, 58542.38, 34063.37, 40876.35, 61315.27]},
    {'name': 'Ross Greetham', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Key', 'currency': 'GBP', 'pr': [26515.13, 31818.45, 47727.42, 30598.75, 36718.5, 55077.76, 32523.53, 39028.25, 58542.37, 34063.37, 40876.35, 61315.27]},
    {'name': 'Lucy Collins', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Key', 'currency': 'GBP', 'pr': [26515.13, 31818.45, 47727.41, 30598.75, 36718.5, 55077.76, 32523.54, 39028.24, 58542.38, 34063.37, 40876.35, 61315.27]},
    {'name': 'George Smith', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Key', 'currency': 'GBP', 'pr': [26515.13, 31818.45, 47727.41, 30598.75, 36718.5, 55077.76, 32523.54, 39028.25, 58542.38, 34063.37, 40876.34, 61315.27]},
    {'name': 'Abbie Lewis', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Key', 'currency': 'GBP', 'pr': [26515.11, 31818.44, 47727.42, 30598.75, 36718.5, 55077.76, 32523.54, 39028.25, 58542.38, 34063.37, 40876.35, 61315.27]},
    {'name': 'Alexa Parritt', 'market': 'United Kingdom', 'team': 'UK SMB Client Sales Key', 'currency': 'GBP', 'pr': [26515.13, 31818.45, 47727.42, 30598.75, 36718.5, 55077.76, 32523.53, 39028.25, 58542.38, 34063.36, 40876.35, 61315.27]},
    {'name': 'Brad Holder', 'market': 'Canada', 'team': 'Canada SMB Client Sales Premier', 'currency': 'CAD', 'pr': [34881.75, 41858.23, 62787.02, 40591.51, 48709.42, 73064.45, 43407.03, 52088.44, 78132.66, 45117.86, 54141.69, 81212.54]},
    {'name': 'Chelsea Salonek', 'market': 'Canada', 'team': 'Canada SMB Client Sales Premier', 'currency': 'USD', 'pr': [34881.75, 41858.23, 62787.02, 40591.51, 48709.42, 73064.45, 43407.03, 52088.44, 78132.66, 45117.86, 54141.69, 81212.54]},
    {'name': 'Carol Murray', 'market': 'Canada', 'team': 'Canada SMB Client Sales Strategic', 'currency': 'CAD', 'pr': [22293.95, 26752.83, 40129.05, 25943.23, 31131.62, 46697.65, 27742.7, 33291.25, 49936.88, 28836.15, 34603.55, 51905.33]},
    {'name': 'Cristian Kawa', 'market': 'Canada', 'team': 'Canada SMB Client Sales Strategic', 'currency': 'CAD', 'pr': [22293.95, 26752.83, 40129.03, 25943.23, 31131.62, 46697.65, 27742.7, 33291.26, 49936.88, 28836.15, 34603.55, 51905.33]},
    {'name': 'Hannah Peach', 'market': 'Canada', 'team': 'Canada SMB Client Sales Strategic', 'currency': 'CAD', 'pr': [22293.95, 26752.83, 40129.03, 25943.24, 31131.62, 46697.65, 27742.7, 33291.25, 49936.88, 28836.15, 34603.55, 51905.33]},
    {'name': 'Tyler Witt', 'market': 'Canada', 'team': 'Canada SMB Client Sales Key', 'currency': 'CAD', 'pr': [13207.5, 15849.04, 23773.46, 15369.42, 18443.16, 27664.86, 16435.48, 19722.57, 29583.87, 17083.27, 20500.02, 30750.03]},
    {'name': 'Deanna Burgess', 'market': 'Canada', 'team': 'Canada SMB Client Sales Key', 'currency': 'CAD', 'pr': [13207.5, 15849.05, 23773.46, 15369.42, 18443.16, 27664.86, 16435.48, 19722.57, 29583.87, 17083.27, 20500.02, 30750.02]},

    {'name': 'Hung Do', 'market': 'Australia', 'team': 'Australia SMB Client Sales', 'currency': 'AUD', 'pr': [23869.06, 28643.23, 42964.41, 28231.22, 33877.38, 50815.86, 29551.99, 35462.64, 53193.75, 30427.99, 36513.51, 54770.05]},
    {'name': 'Manan Taneja', 'market': 'Australia', 'team': 'Australia Mid Market', 'currency': 'AUD', 'pr': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 12102.36, 14522.93, 21784.32, 12461.11, 14953.3, 22429.86]},
]

# Rep → local currency code (derived from QUOTA_DATA, used in the display layer)
REP_CURRENCY: dict[str, str] = {
    row["name"]: row["currency"] for row in QUOTA_DATA
}

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
html,body,[class*="css"]{{font-family:'72','Helvetica Neue',Arial,sans-serif;}}
[data-testid="stSidebar"]{{background:{C['bg_subtle']};border-right:1px solid {C['grey']};}}
.kpi-card{{background:{C['bg']};border:1px solid {C['grey']};border-top:4px solid {C['blue']};
  border-radius:8px;padding:18px 20px 14px;text-align:center;
  box-shadow:0 1px 4px rgba(0,0,0,0.06);}}
.kpi-card.green{{border-top-color:{C['green']};}}
.kpi-card.amber{{border-top-color:{C['amber']};}}
.kpi-card.dark{{border-top-color:{C['dark_blue']};}}
.kpi-value{{font-size:26px;font-weight:700;color:{C['blue']};line-height:1.15;}}
.kpi-value.green{{color:{C['green']};}} .kpi-value.amber{{color:{C['amber']};}}
.kpi-value.dark{{color:{C['dark_blue']};}}
.kpi-label{{font-size:12px;color:{C['dark_grey']};margin-top:5px;
  text-transform:uppercase;letter-spacing:.03em;}}
.sh{{font-size:15px;font-weight:600;color:{C['dark_blue']};
  border-bottom:2px solid {C['blue']};padding-bottom:5px;margin:6px 0 10px;}}
</style>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def normalize(name):
    if pd.isna(name): return name
    s = str(name).strip()
    return NAME_MAP.get(s, s)

def market_from_team(team):
    """Map Concur Team string → Market."""
    if not isinstance(team, str): return "Unknown"
    return TEAM_MARKET.get(team.strip(), "Unknown")

def leader_from_team(team):
    """Map Concur Team string → FLSM name."""
    if not isinstance(team, str): return ""
    return TEAM_LEADER_MAP.get(team.strip(), "")

def director_from_leader(leader):
    """Map FLSM name → Director name."""
    if not isinstance(leader, str): return "Unknown"
    return FLSM_DIRECTOR.get(leader.strip(), "Unknown")

def ltc_rate(term):
    try:    t = int(float(term))
    except: return 0.0
    if t >= 36: return 0.40
    if t >= 24: return 0.30
    if t >= 12: return 0.20
    return 0.0

def bucket(pct):
    if pct is None or (isinstance(pct, float) and np.isnan(pct)): return "No Quota"
    if pct < 0.50:  return "<50%"
    if pct < 0.75:  return "50-75%"
    if pct < 1.00:  return "75-100%"
    if pct < 1.20:  return "100-120%"
    return "120%+"

def kpi_html(label, value, accent=""):
    cls = f"kpi-card {accent}".strip()
    vcls = f"kpi-value {accent}".strip()
    return (f"<div class='{cls}'><div class='{vcls}'>{value}</div>"
            f"<div class='kpi-label'>{label}</div></div>")

def fmt_money(v, sym="$"):
    """Format a monetary value with appropriate scale."""
    if abs(v) >= 1e6: return f"{sym}{v/1e6:.2f}M"
    if abs(v) >= 1e3: return f"{sym}{v/1e3:.0f}K"
    return f"{sym}{v:,.0f}"

def fmt_local(usd_val, currency):
    """Convert a USD amount to the rep's local currency and format for display.

    Backend values are always stored/computed in USD.  This function is used
    only in the display layer (rep-level table) to present amounts in the
    rep's local currency using the 2026 budget exchange rates.
    """
    rate = FX_RATES.get(currency, 1.0)        # local → USD
    local_val = usd_val / rate                 # USD → local
    sym = CURRENCY_SYMBOL.get(currency, "$")
    return fmt_money(local_val, sym)

def fmt_pct(v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return "—"
    return f"{v*100:.1f}%"

def currency_sym(market):
    """Return the currency symbol for KPI cards and charts.

    All market-level and org-level rollups are denominated in USD because
    QUOTA_DATA pr values and SFDC '(converted)' fields are both in USD.
    This function always returns '$'.  Per-rep local-currency display is
    handled separately by fmt_local().
    """
    return "$"

# ─────────────────────────────────────────────────────────────────────────────
# SALESFORCE CONNECTION
# ─────────────────────────────────────────────────────────────────────────────
def sf_connect(username, password, security_token, domain="login"):
    from simple_salesforce import Salesforce
    return Salesforce(username=username, password=password,
                      security_token=security_token, domain=domain)

def sf_connect_session(session_id, instance_url="https://sapconcur.my.salesforce.com"):
    from simple_salesforce import Salesforce
    return Salesforce(session_id=session_id, instance_url=instance_url)

_meta_cache = {}
def _sf_get_meta(sf, report_id):
    if report_id not in _meta_cache:
        _meta_cache[report_id] = sf.restful(
            path=f"analytics/reports/{report_id}",
            method="GET",
            params={"includeDetails": "true"},
            timeout=30,
        )
    return _meta_cache[report_id]

def _cell_val(cell):
    """Extract a plain Python value from a Salesforce Analytics API data cell.

    Handles plain values, currency compound dicts, and User lookup fields.
    For multi-currency orgs the API returns currency amounts as:
      {"value": {"amount": 27467.28, "currency": "GBP"}, "label": "£27,467"}

    Currency normalisation:
      If the compound dict's currency tag is already "USD" (the SFDC org's
      "(converted)" field worked correctly), the amount is returned as-is.
      If the tag is a non-USD currency (e.g. "AED"), SFDC could not convert
      the deal to USD — it returned the local amount instead.  In that case
      we apply FX_RATES to normalise to USD so that the calculation engine
      always operates in a single currency.
    """
    if cell is None:
        return None
    val = cell.get("value")
    if isinstance(val, dict):
        # Currency / compound field: {"amount": X, "currency": "CCY"}
        for key in ("amount", "value", "number"):
            if key in val:
                amount = val[key]
                ccy = val.get("currency", "USD")
                # Normalise to USD if SFDC returned a non-USD local amount.
                if ccy and ccy != "USD":
                    rate = FX_RATES.get(ccy)
                    if rate is not None:
                        try:
                            amount = float(amount) * rate   # local → USD
                        except (TypeError, ValueError):
                            pass
                return amount
        return cell.get("label")
    # Salesforce User / record IDs are 15- or 18-char alphanumeric strings with
    # well-known key prefixes (005 = User, 003 = Contact, 001 = Account, etc.)
    if isinstance(val, str) and len(val) in (15, 18) and val[:3] in (
            "005", "003", "001", "006", "00T", "00U"):
        label = cell.get("label")
        if label:
            return label
    if val is None:
        return cell.get("label")
    return val

# ─────────────────────────────────────────────────────────────────────────────
# SFDC REPORT EXECUTION  (identical to US version)
# ─────────────────────────────────────────────────────────────────────────────
def _post_report(sf, report_id, body_dict, params=None):
    import json as _json
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {sf.session_id}",
    }
    url = f"{sf.base_url}analytics/reports/{report_id}"
    resp = _requests.post(url, headers=headers,
                          json=body_dict, params=params or {}, timeout=120)
    resp.raise_for_status()
    return resp.json()

def _post_report_async(sf, report_id, body_dict, params=None):
    """POST to /instances to get an async instance, then poll until done."""
    import json as _json
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {sf.session_id}",
    }
    inst_url = f"{sf.base_url}analytics/reports/{report_id}/instances"
    resp = _requests.post(inst_url, headers=headers,
                          json=body_dict, params=params or {}, timeout=120)
    resp.raise_for_status()
    inst = resp.json()
    inst_id = inst.get("id") or inst.get("reportId")
    if inst_id is None:
        raise ValueError(f"No instance id in POST /instances response: {inst}")

    # Poll
    for attempt in range(30):
        _time.sleep(1.5 if attempt < 5 else 3)
        poll = sf.restful(
            path=f"analytics/reports/{report_id}/instances/{inst_id}",
            method="GET",
            params=params or {},
            timeout=30,
        )
        if poll.get("status") == "Success":
            return poll, inst_id
        if poll.get("status") in ("Error", "Failed"):
            raise RuntimeError(f"Async report failed: {poll}")
    raise TimeoutError(f"Report instance {inst_id} did not complete in time")

def _fetch_all_pages_async(sf, report_id, inst_id, first_result, base_params):
    rows = list(first_result.get("factMap", {}).get("T!T", {}).get("rows", []))
    if first_result.get("allData", True):
        return rows
    page, MAX = 1, 50
    while page < MAX:
        pg_params = {**base_params, "startRow": page * 2000}
        try:
            r = sf.restful(
                path=f"analytics/reports/{report_id}/instances/{inst_id}",
                method="GET", params=pg_params, timeout=30)
            new = r.get("factMap", {}).get("T!T", {}).get("rows", [])
            if not new: break
            rows.extend(new)
            if r.get("allData", True): break
            page += 1
        except Exception: break
    return rows

def _fetch_all_pages(sf, report_id, first_result, body_dict, base_params):
    rows = list(first_result.get("factMap", {}).get("T!T", {}).get("rows", []))
    if first_result.get("allData", True):
        return rows
    page, MAX = 1, 50
    while page < MAX:
        pg_params = {**base_params, "startRow": page * 2000}
        try:
            if body_dict is not None:
                current = _post_report(sf, report_id, body_dict, pg_params)
            else:
                current = sf.restful(
                    path=f"analytics/reports/{report_id}",
                    method="GET", params=pg_params, timeout=30)
            combined = current.get("factMap", {}).get("T!T", {}).get("rows", [])
            if not combined: break
            rows.extend(combined)
            page += 1
        except Exception: break
    return rows

def sf_run_report(sf, report_id, start_date=None, end_date=None):
    """Fetch a Salesforce Analytics report with a date range filter.

    Strategy (each step is only tried if the previous one fails):
      S1 — synchronous POST with simple CLOSE_DATE filter   (timeout=120s)
      S2 — synchronous POST with full patched metadata       (timeout=120s)
      S3 — async POST with full patched metadata + polling   (timeout=30s/poll)
      G  — GET fallback, no date filter                      (timeout=60s)

    Sync POST (S1/S2) is tried first because it returns in one HTTP round-trip
    with no polling loop, making it far more reliable in environments where
    long-lived connections drop silently.  Async is kept as a fallback for
    reports that explicitly reject sync execution.
    """
    import copy as _copy
    params = {"includeDetails": "true"}
    result = None
    debug  = []
    _winning_body = None

    if "sf_post_debug" not in st.session_state:
        st.session_state.sf_post_debug = {}

    _DATE_COLS = {"CLOSE_DATE","CLOSEDATE","CLOSED_DATE","CREATEDDATE",
                  "LASTMODIFIEDDATE","CLOSE_MONTH","CLOSEMONTH"}

    if start_date and end_date:
        # ── S1: sync POST, simple CLOSE_DATE standard-date-filter ─────────────
        _s1_body = {"reportMetadata": {"standardDateFilter": {
            "column": "CLOSE_DATE", "durationValue": "CUSTOM",
            "startDate": start_date, "endDate": end_date}}}
        try:
            r1 = _post_report(sf, report_id, _s1_body, params)
            n1 = len(r1.get("factMap", {}).get("T!T", {}).get("rows", []))
            result        = r1
            _winning_body = _s1_body
            debug.append(f"S1 OK → {n1} rows")
        except Exception as e1:
            debug.append(f"S1 ERR: {e1}")

        # ── S2: sync POST, full patched metadata ───────────────────────────────
        if result is None:
            try:
                full_resp    = _sf_get_meta(sf, report_id)
                saved_meta   = full_resp.get("reportMetadata", {})
                std_info     = saved_meta.get("standardDateFilter") or {}
                std_col      = std_info.get("column", "CLOSE_DATE")
                patched_meta = _copy.deepcopy(saved_meta)
                patched_meta["standardDateFilter"] = {
                    "column": std_col, "durationValue": "CUSTOM",
                    "startDate": start_date, "endDate": end_date}
                existing_filters = patched_meta.get("reportFilters") or []
                non_date = [f for f in existing_filters
                            if f.get("column","").upper() not in _DATE_COLS]
                patched_meta["reportFilters"] = non_date + [
                    {"column": std_col, "operator": "greaterOrEqual", "value": start_date},
                    {"column": std_col, "operator": "lessOrEqual",    "value": end_date},
                ]
                _s2_body = {"reportMetadata": patched_meta}
                r2 = _post_report(sf, report_id, _s2_body, params)
                n2 = len(r2.get("factMap", {}).get("T!T", {}).get("rows", []))
                result        = r2
                _winning_body = _s2_body
                debug.append(f"S2 OK → {n2} rows")
            except Exception as e2:
                debug.append(f"S2 ERR: {e2}")

        # ── S3: async POST (polling), full patched metadata ───────────────────
        if result is None:
            try:
                _s3_body = _s2_body if "_s2_body" in dir() else _s1_body
                r3, _inst_id = _post_report_async(sf, report_id, _s3_body, params)
                n3 = len(r3.get("factMap", {}).get("T!T", {}).get("rows", []))
                result        = r3
                _winning_body = _s3_body
                debug.append(f"S3-async OK → {n3} rows")
                # async pagination path
                all_rows = _fetch_all_pages_async(sf, report_id, _inst_id,
                                                  result, params)
                st.session_state.sf_post_debug[report_id] = " | ".join(debug)
                meta   = result.get("reportMetadata", {})
                ext    = result.get("reportExtendedMetadata", {})
                cols   = meta.get("detailColumns", [])
                cinfo  = ext.get("detailColumnInfo", {})
                labels = [cinfo.get(c, {}).get("label", c) for c in cols]
                records = []
                for row in all_rows:
                    cells = row.get("dataCells", [])
                    records.append({labels[i]: _cell_val(cells[i])
                                    for i in range(min(len(labels), len(cells)))})
                return pd.DataFrame(records), len(records)
            except Exception as e3:
                debug.append(f"S3-async ERR: {e3}")

    # ── G: GET fallback (no date filter, returns current saved filter) ─────────
    if result is None:
        try:
            result = sf.restful(path=f"analytics/reports/{report_id}",
                                method="GET", params=params, timeout=60)
            nG = len(result.get("factMap", {}).get("T!T", {}).get("rows", []))
            _winning_body = None
            debug.append(f"GET fallback → {nG} rows")
        except Exception as eG:
            raise RuntimeError(
                f"All fetch strategies failed for {report_id}: "
                f"{'; '.join(debug)}; GET: {eG}"
            )

    all_rows = _fetch_all_pages(sf, report_id, result, _winning_body, params)

    st.session_state.sf_post_debug[report_id] = " | ".join(debug)

    meta   = result.get("reportMetadata", {})
    ext    = result.get("reportExtendedMetadata", {})
    cols   = meta.get("detailColumns", [])
    cinfo  = ext.get("detailColumnInfo", {})
    labels = [cinfo.get(c, {}).get("label", c) for c in cols]
    records = []
    for row in all_rows:
        cells = row.get("dataCells", [])
        records.append({labels[i]: _cell_val(cells[i])
                        for i in range(min(len(labels), len(cells)))})
    df = pd.DataFrame(records)
    return df, len(records)


def sf_run_report_multi(sf, report_id, month_nums, year, full_year=False):
    """Run a report for the full YTD period in a single API call.

    Uses a single date range (Jan 1 → end of last month in month_nums) instead
    of one call per month.  The calc engine filters rows by month, so fetching
    the full period at once is both correct and far faster.

    full_year=True: span Jan 1 – Dec 31 (used for Retention, whose date field
    is a text picklist rather than a real date column).
    """
    if isinstance(month_nums, int):
        month_nums = [month_nums]

    start = f"{year}-01-01"
    if full_year:
        end = f"{year}-12-31"
    else:
        last_m    = max(month_nums)
        _, last_d = monthrange(year, last_m)
        end       = f"{year}-{last_m:02d}-{last_d:02d}"

    df, n = sf_run_report(sf, report_id, start, end)
    return df, n


# ─────────────────────────────────────────────────────────────────────────────
# QUOTA LOADING  (INT — uses embedded QUOTA_DATA constant, no file required)
# ─────────────────────────────────────────────────────────────────────────────
def load_quotas_int(month_nums):
    """Build quota maps from the embedded QUOTA_DATA constant.

    PR values in QUOTA_DATA are indexed 0–11 (Jan=0 … Dec=11) and are already
    pro-rated for mid-year starters (zeros before first quota month).

    Returns:
        ldr_quota_map  : FLSM name  → period quota
        rep_quota_map  : rep name   → period quota
        rep_leader_map : rep name   → FLSM name
        rep_market_map : rep name   → market string
        market_pl_map  : market     → total period quota
    """
    if isinstance(month_nums, int):
        month_nums = [month_nums]

    rep_quota_map  = {}
    rep_leader_map = {}
    rep_market_map = {}
    ldr_quota_map  = {}
    market_pl_map  = {"Canada": 0.0, "United Kingdom": 0.0, "Australia": 0.0}

    for row in QUOTA_DATA:
        nm       = normalize(row["name"])
        team     = row["team"]
        market   = row["market"]
        pr       = row["pr"]          # list of 12 monthly values

        # Sum only the selected months (month_nums are 1-based)
        period_q = sum(pr[m - 1] for m in month_nums if 1 <= m <= 12)

        rep_quota_map[nm]  = period_q
        rep_market_map[nm] = market

        leader = TEAM_LEADER_MAP.get(team, "")
        rep_leader_map[nm] = leader

        if leader:
            ldr_quota_map[leader] = ldr_quota_map.get(leader, 0.0) + period_q
        if market in market_pl_map:
            market_pl_map[market] += period_q

    return ldr_quota_map, rep_quota_map, rep_leader_map, rep_market_map, market_pl_map


# ─────────────────────────────────────────────────────────────────────────────
# LEADER ASSIGNMENT HELPER
# ─────────────────────────────────────────────────────────────────────────────
def _assign_leader(df, team_leader_map, mgr_col, rep_leader_map):
    """4-level leader assignment.  Priority (highest → lowest):
      1. Oppty Team → team_leader_map   (INT static map)
      2. Oppty Region → team_leader_map (fallback if region matches a team key)
      3. mgr_col (Opportunity Owner: Manager / Oppty Manager)
      4. rep_leader_map from quota file
    """
    _bad = {"", "nan", "none", "unknown", "n/a"}

    def _clean(s):
        v = normalize(str(s)) if pd.notna(s) else ""
        return "" if str(v).lower().strip() in _bad else str(v).strip()

    # Base: quota file (lowest priority)
    leader = df["Opportunity Owner"].apply(
        lambda r: rep_leader_map.get(_clean(r), "Unknown"))

    # Layer 3: SFDC manager hierarchy
    if mgr_col in df.columns:
        mgr  = df[mgr_col].apply(_clean)
        mask = mgr.ne("")
        leader = leader.where(~mask, mgr)

    if team_leader_map:
        # Layer 2: Oppty Region → team_leader_map
        if "Oppty Region" in df.columns:
            reg_ldr = df["Oppty Region"].apply(
                lambda r: _clean(team_leader_map.get(str(r).strip(), "")))
            mask = reg_ldr.ne("")
            leader = leader.where(~mask, reg_ldr)

        # Layer 1: Oppty Team → team_leader_map (highest priority)
        if "Oppty Team" in df.columns:
            team_ldr = df["Oppty Team"].apply(
                lambda t: _clean(team_leader_map.get(str(t).strip(), "")))
            mask = team_ldr.ne("")
            leader = leader.where(~mask, team_ldr)

    return leader


# ─────────────────────────────────────────────────────────────────────────────
# CALCULATION ENGINE
# ─────────────────────────────────────────────────────────────────────────────
def run_calc(cw_raw, ltc_raw, ret_raw, comp_raw,
             ldr_quota_map, rep_quota_map, rep_leader_map, rep_market_map,
             market_pl_map, month_nums, year):
    """Apply the INT SMB calculation engine.

    Key INT differences vs US:
    - rep_quota_map already holds period quota (no monthly_factor multiplication)
    - Segment → Market grouping
    - Leader assignment uses TEAM_LEADER_MAP (static) as primary source
    """
    if isinstance(month_nums, int):
        month_nums = [month_nums]
    month_names = [MONTH_NAMES[m - 1] for m in month_nums]

    # ── Filter by period ─────────────────────────────────────────────────────
    cw = cw_raw.copy()
    cw["Close Date"] = pd.to_datetime(cw["Close Date"], errors="coerce")
    cw = cw[(cw["Close Date"].dt.month.isin(month_nums)) &
            (cw["Close Date"].dt.year  == year)].copy()
    cw["Opportunity Owner"] = cw["Opportunity Owner"].apply(normalize)
    mgr_col_cw = ("Oppty Manager" if "Oppty Manager" in cw.columns
                  else "Opportunity Owner: Manager")
    if mgr_col_cw in cw.columns:
        cw[mgr_col_cw] = cw[mgr_col_cw].apply(normalize)

    ltc = ltc_raw.copy()
    ltc["Close Date"] = pd.to_datetime(ltc["Close Date"], errors="coerce")
    ltc = ltc[(ltc["Close Date"].dt.month.isin(month_nums)) &
              (ltc["Close Date"].dt.year  == year)].copy()
    ltc["Opportunity Owner"] = ltc["Opportunity Owner"].apply(normalize)
    mgr_col_ltc = ("Oppty Manager" if "Oppty Manager" in ltc.columns
                   else "Opportunity Owner: Manager")
    if mgr_col_ltc in ltc.columns:
        ltc[mgr_col_ltc] = ltc[mgr_col_ltc].apply(normalize)

    # ── Retention — year-aware FMC filter ────────────────────────────────────
    ret = ret_raw.copy()
    _fmc = ret["Final Month Closed"].fillna("").astype(str).str.strip()
    _month_year_explicit = [f"{mn} {year}"       for mn in month_names]
    _month_year_short    = [f"{mn[:3]} {year}"   for mn in month_names]
    _month_year_hyphen   = [f"{mn[:3]}-{year}"   for mn in month_names]
    _month_only          = month_names
    _match_explicit = _fmc.isin(_month_year_explicit + _month_year_short +
                                 _month_year_hyphen)
    if "Close Date" in ret.columns:
        _close_yr = pd.to_datetime(ret["Close Date"], errors="coerce").dt.year
        _match_month_only = _fmc.isin(_month_only) & _close_yr.eq(year).fillna(False)
    else:
        _match_month_only = _fmc.isin(_month_only)
    ret = ret[_match_explicit | _match_month_only].copy()
    ret["Opportunity Owner"] = ret["Opportunity Owner"].apply(normalize)
    mgr_col_ret = ("Oppty Manager" if "Oppty Manager" in ret.columns
                   else "Opportunity Owner: Manager")
    if mgr_col_ret in ret.columns:
        ret[mgr_col_ret] = ret[mgr_col_ret].apply(normalize)
    ret["BMI_ARR"]     = pd.to_numeric(ret["BMI Sales ARR"], errors="coerce").fillna(0)
    # Roll-up Sales Credit — try converted column first, fall back to plain column
    _sc_col = ("Roll-up Sales Credit Calculation (converted)"
               if "Roll-up Sales Credit Calculation (converted)" in ret.columns
               else "Roll-up Sales Credit Calculation"
               if "Roll-up Sales Credit Calculation" in ret.columns
               else None)
    ret["Roll_SC_Ret"] = (pd.to_numeric(ret[_sc_col], errors="coerce").fillna(0)
                          if _sc_col else 0.0)
    # Exclude Split Opportunity rows
    _split = ret.get("ARR Disputes", pd.Series([""] * len(ret))).fillna("")
    ret = ret[~_split.str.contains("Split Opportunity", case=False, na=False)].copy()

    # ── Complete ──────────────────────────────────────────────────────────────
    comp = comp_raw.copy()
    comp["_cm_parsed"] = pd.to_datetime(
        comp["Close Month"].astype(str).str.strip(), errors="coerce", dayfirst=False)
    comp = comp[(comp["_cm_parsed"].dt.year  == year) &
                (comp["_cm_parsed"].dt.month.isin(month_nums))].copy()
    comp.drop(columns=["_cm_parsed"], inplace=True)
    comp["Opportunity Owner"] = comp["Opportunity Owner"].apply(normalize)
    _sc_comp = ("Roll-up Sales Credit Calculation (converted)"
                if "Roll-up Sales Credit Calculation (converted)" in comp.columns
                else "Roll-up Sales Credit Calculation"
                if "Roll-up Sales Credit Calculation" in comp.columns
                else None)
    comp["Complete_Credit_Val"] = (pd.to_numeric(comp[_sc_comp], errors="coerce").fillna(0)
                                   if _sc_comp else 0.0)

    # ── Lookup tables ─────────────────────────────────────────────────────────
    complete_names  = set(comp["Opportunity Name"].str.strip())
    complete_lookup = dict(zip(comp["Opportunity Name"].str.strip(),
                               comp["Complete_Credit_Val"]))

    _fa_ltc = ("Forecast Amount (converted)" if "Forecast Amount (converted)" in ltc.columns
               else "Forecast Amount")
    ltc["LTC_Uplift_Calc"] = ltc.apply(
        lambda r: pd.to_numeric(r[_fa_ltc], errors="coerce")
                  * ltc_rate(r["Term (no. of months)"]), axis=1)
    ltc_lookup = dict(zip(ltc["Opportunity Name"].str.strip(),
                          ltc["LTC_Uplift_Calc"]))

    # ── Retention grouping ───────────────────────────────────────────────────
    _team_col_ret = ("Oppty Team" if "Oppty Team" in ret.columns
                     else "Opportunity Owner" if True else None)
    ret_grp_agg = {
        "BMI_SUM":  ("BMI_ARR",            "sum"),
        "SC_SUM":   ("Roll_SC_Ret",         "sum"),
        "Rep":      ("Opportunity Owner",   "first"),
        "OppName":  ("Opportunity Name",    "first"),
    }
    if "Oppty Team" in ret.columns:
        ret_grp_agg["Team"] = ("Oppty Team", "first")
    ret_grp = ret.groupby("Opportunity ID").agg(**ret_grp_agg).reset_index()
    ret_grp["Retention_Credit"] = (ret_grp["BMI_SUM"] - ret_grp["SC_SUM"]).clip(lower=0)

    # Assign leader to retention deals
    _rg = ret_grp.rename(columns={"Team": "Oppty Team"} if "Team" in ret_grp.columns else {})
    _rg["Opportunity Owner"] = _rg["Rep"]
    if mgr_col_ret in ret.columns:
        _first_mgr = (ret.groupby("Opportunity ID")[mgr_col_ret]
                      .first().reset_index()
                      .rename(columns={mgr_col_ret: "_mgr_ret"}))
        _rg = _rg.merge(_first_mgr, on="Opportunity ID", how="left")
        _rg[mgr_col_ret] = _rg.get("_mgr_ret", "")
    else:
        _rg[mgr_col_ret] = ""
    ret_grp["Leader"] = _assign_leader(_rg, TEAM_LEADER_MAP,
                                        mgr_col_ret, rep_leader_map).values

    ret_by_oppname = ret_grp.groupby("OppName")["Retention_Credit"].sum().to_dict()

    # ── Master CW dataset ─────────────────────────────────────────────────────
    master = cw.copy()
    master["Rep"] = master["Opportunity Owner"]
    master["Leader"] = _assign_leader(master, TEAM_LEADER_MAP,
                                       mgr_col_cw, rep_leader_map)

    # Team / Market columns
    team_col = ("Oppty Team" if "Oppty Team" in master.columns
                else "Owner Team" if "Owner Team" in master.columns
                else None)
    if team_col:
        master["Team"] = master[team_col].fillna("Unknown").str.strip()
    else:
        master["Team"] = "Unknown"
    master["Market"] = master["Team"].apply(market_from_team)
    master["Director"] = master["Leader"].apply(director_from_leader)

    _fa_cw = ("Forecast Amount (converted)" if "Forecast Amount (converted)" in master.columns
              else "Forecast Amount")
    master["Forecast_Amount_ARR"] = pd.to_numeric(
        master[_fa_cw], errors="coerce").fillna(0)
    master["_OppName"]       = master["Opportunity Name"].str.strip()
    master["In_Complete"]    = master["_OppName"].isin(complete_names).astype(int)
    master["Complete_Credit"]  = master["_OppName"].map(complete_lookup).fillna(0)
    master["CW_ARR_Adjusted"]  = np.where(
        master["In_Complete"] == 1, 0, master["Forecast_Amount_ARR"])
    master["LTC_Uplift"]       = master["_OppName"].map(ltc_lookup).fillna(0)
    master["Retention_Credit"] = master["_OppName"].map(ret_by_oppname).fillna(0)
    master["Total_Credited"]   = (master["CW_ARR_Adjusted"] + master["LTC_Uplift"]
                                   + master["Retention_Credit"]
                                   + master["Complete_Credit"])

    # ── Standalone retention (opps not matched to CW) ────────────────────────
    cw_names   = set(master["_OppName"])
    total_ret  = ret_grp["Retention_Credit"].sum()

    # ── Rep aggregation ───────────────────────────────────────────────────────
    master_s = master.sort_values("Close Date", na_position="first")
    cw_by_rep = master_s.groupby("Rep").agg(
        Leader_cw       = ("Leader",          "last"),
        Team_cw         = ("Team",            "last"),
        Market_cw       = ("Market",          "last"),
        Director_cw     = ("Director",        "last"),
        CW_ARR          = ("CW_ARR_Adjusted", "sum"),
        LTC_Credit      = ("LTC_Uplift",      "sum"),
        Complete_Credit = ("Complete_Credit", "sum"),
        CW_Units        = ("Opportunity Name","count"),
    ).reset_index()

    ret_by_rep = ret_grp.groupby("Rep").agg(
        Retention_Credit = ("Retention_Credit", "sum"),
        Leader_ret       = ("Leader",           "first"),
    ).reset_index()

    rep_grp = cw_by_rep.merge(
        ret_by_rep[["Rep","Retention_Credit","Leader_ret"]],
        on="Rep", how="outer")
    for col in ["CW_ARR","LTC_Credit","Complete_Credit","CW_Units","Retention_Credit"]:
        rep_grp[col] = rep_grp[col].fillna(0)

    def resolve_leader(row):
        for src in [row.get("Leader_cw",""), row.get("Leader_ret","")]:
            if isinstance(src, str) and src.strip().lower() not in ("","nan","unknown","n/a"):
                return src.strip()
        return rep_leader_map.get(row["Rep"], "Unknown")

    def resolve_team(row):
        for src in [row.get("Team_cw","")]:
            if isinstance(src, str) and src.strip().lower() not in ("","nan","unknown"):
                return src.strip()
        return "Unknown"

    def resolve_market(row):
        t = resolve_team(row)
        m = market_from_team(t)
        if m != "Unknown": return m
        # Fall back to rep_market_map from quota file
        return rep_market_map.get(row["Rep"], "Unknown")

    rep_grp["Leader"]   = rep_grp.apply(resolve_leader,  axis=1)
    rep_grp["Team"]     = rep_grp.apply(resolve_team,    axis=1)
    rep_grp["Market"]   = rep_grp.apply(resolve_market,  axis=1)
    rep_grp["Director"] = rep_grp["Leader"].apply(director_from_leader)

    for _c in ["CW_ARR","LTC_Credit","Retention_Credit","Complete_Credit","CW_Units"]:
        rep_grp[_c] = pd.to_numeric(rep_grp[_c], errors="coerce").fillna(0)
    rep_grp["Total_Credited"] = (rep_grp["CW_ARR"] + rep_grp["LTC_Credit"]
                                  + rep_grp["Retention_Credit"]
                                  + rep_grp["Complete_Credit"])
    # Period quota: already a dollar amount (no monthly_factor needed)
    rep_grp["Period_Quota"] = pd.to_numeric(
        rep_grp["Rep"].map(lambda r: rep_quota_map.get(r, 0)),
        errors="coerce").fillna(0)
    rep_grp["Pct_to_Quota"] = np.where(
        rep_grp["Period_Quota"] > 0,
        rep_grp["Total_Credited"] / rep_grp["Period_Quota"], np.nan)
    rep_grp = rep_grp[["Rep","Leader","Director","Team","Market",
                        "CW_ARR","LTC_Credit","Retention_Credit","Complete_Credit",
                        "Total_Credited","Period_Quota","Pct_to_Quota","CW_Units"]]

    # ── Leader aggregation ────────────────────────────────────────────────────
    ldr_grp = rep_grp.groupby("Leader").agg(
        Director        = ("Director",         "first"),
        Market          = ("Market",           "first"),
        CW_ARR          = ("CW_ARR",           "sum"),
        LTC_Credit      = ("LTC_Credit",       "sum"),
        Retention_Credit= ("Retention_Credit", "sum"),
        Complete_Credit = ("Complete_Credit",  "sum"),
        Total_Credited  = ("Total_Credited",   "sum"),
        CW_Units        = ("CW_Units",         "sum"),
    ).reset_index()
    ldr_grp["Period_Quota"] = pd.to_numeric(
        ldr_grp["Leader"].map(lambda l: ldr_quota_map.get(l, 0)),
        errors="coerce").fillna(0)
    ldr_grp["Pct_to_Quota"] = np.where(
        ldr_grp["Period_Quota"] > 0,
        ldr_grp["Total_Credited"] / ldr_grp["Period_Quota"], np.nan)

    # ── Market aggregation ────────────────────────────────────────────────────
    mkt_grp = rep_grp[rep_grp["Market"].isin(["Canada","United Kingdom","Australia"])
                      ].groupby("Market").agg(
        CW_ARR          = ("CW_ARR",           "sum"),
        LTC_Credit      = ("LTC_Credit",       "sum"),
        Retention_Credit= ("Retention_Credit", "sum"),
        Complete_Credit = ("Complete_Credit",  "sum"),
        Total_Credited  = ("Total_Credited",   "sum"),
        CW_Units        = ("CW_Units",         "sum"),
    ).reset_index()
    mkt_grp["PL_Quota"] = pd.to_numeric(
        mkt_grp["Market"].map(market_pl_map), errors="coerce").fillna(0)
    mkt_grp["Pct_to_PL"] = np.where(
        mkt_grp["PL_Quota"] > 0,
        mkt_grp["Total_Credited"] / mkt_grp["PL_Quota"], np.nan)

    # ── Attainment distribution ───────────────────────────────────────────────
    rq = rep_grp[rep_grp["Period_Quota"] > 0].copy()
    rq["Bucket"] = rq["Pct_to_Quota"].apply(bucket)
    dist = rq.groupby("Bucket").size().reset_index(name="Count")

    # ── Org summary ───────────────────────────────────────────────────────────
    org_total = rep_grp["Total_Credited"].sum()
    org_quota = sum(market_pl_map.values())
    org = {
        "Total_Credited":   org_total,
        "CW_ARR":           rep_grp["CW_ARR"].sum(),
        "LTC_Credit":       rep_grp["LTC_Credit"].sum(),
        "Retention_Credit": total_ret,
        "Complete_Credit":  rep_grp["Complete_Credit"].sum(),
        "PL_Quota":         org_quota,
        "Pct_to_PL":        org_total / org_quota if org_quota else 0,
        "CW_Units":         int(len(cw)),
        "Above_100":        int((rq["Pct_to_Quota"] >= 1.0).sum()),
        "Above_120":        int((rq["Pct_to_Quota"] >= 1.2).sum()),
        "Total_Reps":       int(len(rq)),
        "Row_Counts": {"CW": len(cw), "LTC": len(ltc),
                       "Retention": len(ret), "Complete": len(comp)},
    }

    return {"org": org, "market": mkt_grp, "leader": ldr_grp,
            "rep": rep_grp, "dist": dist, "master": master}


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────────────────────────────────────
for k, v in [("sf", None), ("data", None), ("last_run", None),
              ("raw_cw", None), ("raw_ltc", None), ("raw_ret", None),
              ("raw_comp", None)]:
    if k not in st.session_state:
        st.session_state[k] = v

period_label = "—"
sel_market   = "All Markets"

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### SAP Concur")
    st.markdown("**INT SMB Client Sales Dashboard**")
    st.markdown("---")

    # ── Credentials ──────────────────────────────────────────────────────────
    with st.expander("Salesforce Connection", expanded=st.session_state.sf is None):
        secrets_sf = st.secrets.get("salesforce", {}) if hasattr(st, "secrets") else {}

        auth_mode = st.radio("Auth Method",
                             ["Session ID (SSO / Microsoft login)", "Username + Password"],
                             index=0, horizontal=True)

        if auth_mode == "Session ID (SSO / Microsoft login)":
            session_id = st.text_input("Session ID", type="password",
                                       value=secrets_sf.get("session_id", ""),
                                       placeholder="00D…  (paste from browser — see instructions below)",
                                       help="Log into Salesforce in your browser → F12 → "
                                            "Application → Cookies → sapconcur.my.salesforce.com → "
                                            "copy the value of the 'sid' cookie")
            instance_url = st.text_input("Instance URL",
                                         value=secrets_sf.get("instance_url",
                                                               "https://sapconcur.my.salesforce.com"))
            col_a, col_b = st.columns(2)
            if col_a.button("Connect", use_container_width=True):
                if not session_id:
                    st.error("Paste your Session ID first.")
                else:
                    try:
                        st.session_state.sf = sf_connect_session(session_id, instance_url)
                        st.success("Connected")
                    except Exception as e:
                        st.error(str(e))
        else:
            username  = st.text_input("Username (email)",
                                      value=secrets_sf.get("username",""),
                                      placeholder="you@company.com")
            password  = st.text_input("Password", type="password",
                                      value=secrets_sf.get("password",""))
            sec_token = st.text_input("Security Token", type="password",
                                      value=secrets_sf.get("security_token",""))
            domain    = st.text_input("Domain",
                                      value=secrets_sf.get("domain","login"))
            col_a, col_b = st.columns(2)
            if col_a.button("Connect", use_container_width=True):
                try:
                    st.session_state.sf = sf_connect(username, password, sec_token, domain)
                    st.success("Connected")
                except Exception as e:
                    st.error(str(e))

        if col_b.button("Disconnect", use_container_width=True):
            st.session_state.sf = None
            st.session_state.data = None

    conn_ok = st.session_state.sf is not None
    st.markdown(
        f"<span style='color:{'#188918' if conn_ok else '#BB0000'};font-weight:600'>"
        f"{'● Connected' if conn_ok else '○ Not connected'}</span>",
        unsafe_allow_html=True)

    st.markdown("---")

    # ── Report IDs ────────────────────────────────────────────────────────────
    with st.expander("Report IDs", expanded=True):
        secrets_rpt = st.secrets.get("reports", {}) if hasattr(st, "secrets") else {}
        rpt_cw   = st.text_input("CW ARR Report ID",
                                  value=secrets_rpt.get("cw_arr_report_id",   "00OPg00000Qf2dB"),
                                  placeholder="00O…")
        rpt_ltc  = st.text_input("LTC Report ID",
                                  value=secrets_rpt.get("ltc_report_id",      "00OPg00000Qf2i1"),
                                  placeholder="00O…")
        rpt_ret  = st.text_input("Retention Report ID",
                                  value=secrets_rpt.get("retention_report_id","00OPg00000Qf39R"),
                                  placeholder="00O…")
        rpt_comp = st.text_input("Complete / Referral Report ID",
                                  value=secrets_rpt.get("complete_report_id", "00OPg00000Qf2mr"),
                                  placeholder="00O…")

    st.markdown("---")

    # ── Reporting Period ──────────────────────────────────────────────────────
    st.markdown("**Reporting Period**")
    period_type = st.radio("Period Type", ["Monthly","Quarterly","YTD"],
                           horizontal=True)
    sel_year = st.number_input("Year", min_value=2020, max_value=2030,
                                value=2026, step=1)
    QUARTER_MONTHS = {"Q1":[1,2,3],"Q2":[4,5,6],"Q3":[7,8,9],"Q4":[10,11,12]}

    if period_type == "Monthly":
        sel_month    = st.selectbox("Month", MONTH_NAMES,
                                    index=MONTH_NAMES.index("July"))
        month_nums   = [MONTH_NAMES.index(sel_month) + 1]
        period_label = f"{sel_month} {int(sel_year)}"
    elif period_type == "Quarterly":
        sel_quarter  = st.selectbox("Quarter", ["Q1","Q2","Q3","Q4"], index=2)
        month_nums   = QUARTER_MONTHS[sel_quarter]
        period_label = f"{sel_quarter} {int(sel_year)}"
    else:
        sel_month    = st.selectbox("Through Month", MONTH_NAMES,
                                    index=MONTH_NAMES.index("July"))
        month_nums   = list(range(1, MONTH_NAMES.index(sel_month) + 2))
        period_label = f"YTD through {sel_month} {int(sel_year)}"

    st.markdown("---")

    # ── Run Reports ───────────────────────────────────────────────────────────
    run_btn = st.button("Run Reports", type="primary",
                        use_container_width=True, disabled=not conn_ok)
    if run_btn:
        if not all([rpt_cw, rpt_ltc, rpt_ret, rpt_comp]):
            st.error("All 4 Report IDs are required.")
        else:
            n_months = len(month_nums)
            sf   = st.session_state.sf
            _yr  = int(sel_year)
            try:
                with st.spinner("Fetching CW ARR report…"):
                    cw_df,  n1 = sf_run_report_multi(sf, rpt_cw,  month_nums, _yr)
                with st.spinner("Fetching LTC report…"):
                    ltc_df, n2 = sf_run_report_multi(sf, rpt_ltc, month_nums, _yr)
                with st.spinner("Fetching Retention report…"):
                    ret_df, n3 = sf_run_report_multi(sf, rpt_ret, month_nums, _yr,
                                                     full_year=True)
                with st.spinner("Fetching Referral report…"):
                    comp_df,n4 = sf_run_report_multi(sf, rpt_comp,month_nums, _yr)

                st.session_state.raw_cw   = cw_df
                st.session_state.raw_ltc  = ltc_df
                st.session_state.raw_ret  = ret_df
                st.session_state.raw_comp = comp_df
                st.caption(f"CW ARR: {n1} rows | LTC: {n2} rows | "
                           f"Retention: {n3} rows | Referral: {n4} rows")
                for name, n in [("CW ARR",n1),("LTC",n2),
                                ("Retention",n3),("Referral",n4)]:
                    if n == 0:
                        st.warning(f"**{name} returned 0 rows.** "
                                   f"Check the report date filter in Salesforce.")
                    elif n >= 2000:
                        st.warning(f"**{name} returned 2,000 rows** — API cap. "
                                   f"Some records may be missing.")
            except Exception as e:
                st.error(f"Report error: {e}")
                st.stop()

            with st.spinner("Calculating metrics…"):
                try:
                    lq, rq, rlm, rmm, mpl = load_quotas_int(month_nums)
                    result = run_calc(
                        st.session_state.raw_cw,  st.session_state.raw_ltc,
                        st.session_state.raw_ret, st.session_state.raw_comp,
                        lq, rq, rlm, rmm, mpl, month_nums, int(sel_year))
                    st.session_state.data     = result
                    st.session_state.last_run = datetime.now()
                except Exception as e:
                    st.error(f"Calculation error: {e}\n\n"
                             f"Check the column preview in the diagnostic expander.")
                    st.stop()
            st.rerun()

    # ── Filters (shown after data loads) ─────────────────────────────────────
    if st.session_state.data:
        st.markdown("---")
        st.markdown("**Filters**")
        rep_df_all = st.session_state.data["rep"]

        mkts_avail = ["All Markets"] + sorted(
            rep_df_all[rep_df_all["Market"].isin(
                ["Canada","United Kingdom","Australia"])]["Market"].unique())
        sel_market = st.selectbox("Market", mkts_avail)

        _rep_mkt = (rep_df_all[rep_df_all["Market"] == sel_market]
                    if sel_market != "All Markets" else rep_df_all)
        teams_avail = ["All Teams"] + sorted([
            t for t in _rep_mkt["Team"].unique()
            if isinstance(t, str) and t.strip() not in ("", "Unknown", "nan")
        ])
        sel_team = st.selectbox("Team", teams_avail)

        _rep_team = (_rep_mkt[_rep_mkt["Team"] == sel_team]
                     if sel_team != "All Teams" else _rep_mkt)
        ldrs_avail = ["All Leaders"] + sorted(
            _rep_team["Leader"].dropna().unique().tolist())
        sel_ldr = st.selectbox("Leader", ldrs_avail)
    else:
        sel_market = "All Markets"
        sel_team   = "All Teams"
        sel_ldr    = "All Leaders"

    if st.session_state.last_run:
        st.caption(f"Last run: {st.session_state.last_run.strftime('%b %d %Y %H:%M')}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN — SPLASH
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state.data is None:
    st.markdown(
        f"<div style='text-align:center;padding:80px 0'>"
        f"<h1 style='color:{C['blue']};font-size:2rem;margin-bottom:.5rem'>"
        f"SAP Concur INT SMB Client Sales Dashboard</h1>"
        f"<p style='color:{C['dark_grey']};font-size:1rem'>"
        f"Markets: Canada (CAD) &nbsp;·&nbsp; United Kingdom (GBP) &nbsp;·&nbsp; Australia (AUD)</p>"
        f"<p style='color:{C['dark_grey']};font-size:.95rem;margin-top:20px'>"
        f"Connect to Salesforce and run reports using the sidebar to begin.</p>"
        f"</div>", unsafe_allow_html=True)
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# DATA + FILTER HELPERS
# ─────────────────────────────────────────────────────────────────────────────
d        = st.session_state.data
org_d    = d["org"]
rep_df   = d["rep"].copy()
ldr_df   = d["leader"].copy()
mkt_df   = d["market"].copy()
dist_df  = d["dist"].copy()

# Apply market / team / leader filters
if sel_market != "All Markets":
    rep_df  = rep_df[rep_df["Market"] == sel_market].copy()
    ldr_df  = ldr_df[ldr_df["Market"] == sel_market].copy()
    mkt_df  = mkt_df[mkt_df["Market"] == sel_market].copy()

if "sel_team" in dir() and sel_team != "All Teams":
    rep_df = rep_df[rep_df["Team"] == sel_team].copy()
    ldr_df = ldr_df[ldr_df["Leader"].isin(
        [TEAM_LEADER_MAP.get(sel_team, "")])].copy()

if "sel_ldr" in dir() and sel_ldr != "All Leaders":
    rep_df = rep_df[rep_df["Leader"] == sel_ldr].copy()
    ldr_df = ldr_df[ldr_df["Leader"] == sel_ldr].copy()

# All KPI cards, charts, and rollups are in USD.
# Per-rep local currency display is applied in the Rep Performance table below.
curr_sym = "$"
fmt_m = lambda v: fmt_money(v, "$")

# Filtered totals
f_total    = rep_df["Total_Credited"].sum()
f_quota    = rep_df["Period_Quota"].sum()
f_cw       = rep_df["CW_ARR"].sum()
f_ltc      = rep_df["LTC_Credit"].sum()
f_ret      = rep_df["Retention_Credit"].sum()
f_comp     = rep_df["Complete_Credit"].sum()
f_units    = int(rep_df["CW_Units"].sum())
f_pct      = f_total / f_quota if f_quota else 0.0

# Rep attainment distribution (filtered)
rq_filt    = rep_df[rep_df["Period_Quota"] > 0].copy()
rq_filt["Bucket"] = rq_filt["Pct_to_Quota"].apply(bucket)
dist_filt  = rq_filt.groupby("Bucket").size().reset_index(name="Count")
above_100  = int((rq_filt["Pct_to_Quota"] >= 1.0).sum())
above_120  = int((rq_filt["Pct_to_Quota"] >= 1.2).sum())
total_reps = len(rq_filt)


# ─────────────────────────────────────────────────────────────────────────────
# CURRENCY NOTE
# ─────────────────────────────────────────────────────────────────────────────
# All quota values and SFDC "(converted)" performance fields are in USD.
# KPI cards, charts, and market/leader rollups are shown in USD ($).
# The Rep Performance table converts each rep's figures to their local
# currency using 2026 budget exchange rates (FX_RATES constant).

# ─────────────────────────────────────────────────────────────────────────────
# TITLE ROW
# ─────────────────────────────────────────────────────────────────────────────
mkt_label = (f" — {sel_market}"
             if sel_market != "All Markets" else " — All Markets")
st.markdown(
    f"<h2 style='color:{C['dark_blue']};margin-bottom:4px'>"
    f"INT SMB Client Sales{mkt_label}</h2>"
    f"<p style='color:{C['dark_grey']};margin-top:0'>{period_label}"
    f" &nbsp;·&nbsp; Rollups in USD &nbsp;·&nbsp; Rep detail in local currency</p>",
    unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# ROW 1 — KPI CARDS
# ─────────────────────────────────────────────────────────────────────────────
c1, c2, c3, c4, c5, c6 = st.columns(6)
pct_color = ("green" if f_pct >= 1.0 else "amber" if f_pct >= 0.75 else "")

c1.markdown(kpi_html("Total Credited",    fmt_m(f_total),   pct_color), unsafe_allow_html=True)
c2.markdown(kpi_html("Period Quota",      fmt_m(f_quota),   "dark"),    unsafe_allow_html=True)
c3.markdown(kpi_html("% to Quota",        fmt_pct(f_pct),   pct_color), unsafe_allow_html=True)
c4.markdown(kpi_html("CW Units",          str(f_units),     ""),        unsafe_allow_html=True)
c5.markdown(kpi_html("Reps ≥ 100%",
                      f"{above_100} / {total_reps}",          "green"),  unsafe_allow_html=True)
c6.markdown(kpi_html("Reps ≥ 120%",
                      f"{above_120} / {total_reps}",          "green"),  unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# ROW 2 — MARKET PERFORMANCE + CREDIT BREAKDOWN
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='sh'>Market Performance</div>", unsafe_allow_html=True)
col_mkt, col_pie = st.columns([3, 2])

with col_mkt:
    if len(mkt_df) == 0:
        st.info("No market data for the current filter.")
    else:
        fig_mkt = go.Figure()
        # Stacked bars: CW ARR, LTC, Retention, Complete
        components = [
            ("CW ARR",         "CW_ARR",           C["blue"]),
            ("LTC Uplift",     "LTC_Credit",        C["light_blue"]),
            ("Retention",      "Retention_Credit",  C["dark_blue"]),
            ("Complete/Ref",   "Complete_Credit",   C["dark_grey"]),
        ]
        for lbl, col_name, clr in components:
            fig_mkt.add_trace(go.Bar(
                name=lbl, x=mkt_df["Market"], y=mkt_df[col_name],
                marker_color=clr, text=mkt_df[col_name].apply(
                    lambda v: f"{curr_sym}{v/1e3:.0f}K" if v >= 1000 else ""),
                textposition="inside"))

        # Quota line
        fig_mkt.add_trace(go.Scatter(
            name="Quota", x=mkt_df["Market"], y=mkt_df["PL_Quota"],
            mode="markers+lines",
            marker=dict(size=10, symbol="diamond", color=C["amber"]),
            line=dict(color=C["amber"], width=2, dash="dot")))

        fig_mkt.update_layout(
            barmode="stack", height=300, margin=dict(l=0, r=0, t=30, b=0),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
            xaxis_title="", yaxis_title="",
            plot_bgcolor=C["bg"], paper_bgcolor=C["bg"],
            font=dict(family="72,Helvetica Neue,Arial,sans-serif", size=12))
        fig_mkt.update_yaxes(tickprefix=curr_sym if curr_sym else "",
                              showgrid=True, gridcolor=C["grey"])
        st.plotly_chart(fig_mkt, use_container_width=True)

with col_pie:
    labels_pie = ["CW ARR", "LTC Uplift", "Retention", "Complete/Ref"]
    vals_pie   = [f_cw, f_ltc, f_ret, f_comp]
    vals_pie   = [max(0, v) for v in vals_pie]
    if sum(vals_pie) > 0:
        fig_pie = go.Figure(go.Pie(
            labels=labels_pie, values=vals_pie, hole=0.55,
            marker_colors=[C["blue"], C["light_blue"], C["dark_blue"], C["dark_grey"]],
            textinfo="percent", textfont_size=11))
        fig_pie.update_layout(
            height=300, margin=dict(l=0, r=0, t=30, b=0),
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.15, x=0.5,
                        xanchor="center", font=dict(size=11)),
            plot_bgcolor=C["bg"], paper_bgcolor=C["bg"],
            annotations=[dict(text=f"{fmt_m(f_total)}<br><span style='font-size:9px'>"
                                   f"Total</span>",
                              font_size=13, showarrow=False)])
        st.plotly_chart(fig_pie, use_container_width=True)
    else:
        st.info("No credited amounts for current filter.")


# ─────────────────────────────────────────────────────────────────────────────
# ROW 3 — LEADER ATTAINMENT + DISTRIBUTION
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='sh'>Leader Attainment</div>", unsafe_allow_html=True)
col_ldr, col_dist = st.columns([3, 2])

with col_ldr:
    ldr_show = ldr_df[ldr_df["Leader"].ne("Unknown")].copy()
    if len(ldr_show) == 0:
        st.info("No leader data for the current filter.")
    else:
        ldr_show = ldr_show.sort_values("Total_Credited", ascending=True)
        ldr_show["Bar_Color"] = ldr_show["Market"].map(
            lambda m: MARKET_COLORS.get(m, C["blue"]))
        ldr_show["Pct_Label"] = ldr_show["Pct_to_Quota"].apply(
            lambda p: f"  {p*100:.0f}%" if pd.notna(p) else "")
        fig_ldr = go.Figure()
        for mkt, clr in MARKET_COLORS.items():
            sub = ldr_show[ldr_show["Market"] == mkt]
            if len(sub) == 0: continue
            fig_ldr.add_trace(go.Bar(
                name=mkt, y=sub["Leader"], x=sub["Total_Credited"],
                orientation="h", marker_color=clr,
                text=sub["Pct_Label"], textposition="outside"))
        # Quota markers
        ldr_q = ldr_show[ldr_show["Period_Quota"] > 0]
        fig_ldr.add_trace(go.Scatter(
            name="Quota", y=ldr_q["Leader"], x=ldr_q["Period_Quota"],
            mode="markers",
            marker=dict(size=8, symbol="line-ns", color=C["amber"],
                        line=dict(width=2, color=C["amber"]))))
        fig_ldr.update_layout(
            barmode="overlay", height=max(280, len(ldr_show) * 38),
            margin=dict(l=0, r=60, t=20, b=0),
            legend=dict(orientation="h", y=1.06, x=0),
            xaxis_title="", yaxis_title="",
            plot_bgcolor=C["bg"], paper_bgcolor=C["bg"],
            font=dict(family="72,Helvetica Neue,Arial,sans-serif", size=12))
        fig_ldr.update_xaxes(tickprefix=curr_sym if curr_sym else "",
                              showgrid=True, gridcolor=C["grey"])
        st.plotly_chart(fig_ldr, use_container_width=True)

with col_dist:
    if len(dist_filt) > 0:
        dist_sorted = pd.DataFrame({"Bucket": BUCKET_ORDER}).merge(
            dist_filt, on="Bucket", how="left").fillna(0)
        dist_sorted["Count"] = dist_sorted["Count"].astype(int)
        fig_dist = go.Figure(go.Bar(
            x=dist_sorted["Bucket"], y=dist_sorted["Count"],
            marker_color=[BUCKET_COLORS.get(b, C["grey"]) for b in dist_sorted["Bucket"]],
            text=dist_sorted["Count"], textposition="outside"))
        fig_dist.update_layout(
            height=280, margin=dict(l=0, r=0, t=30, b=0),
            xaxis_title="Attainment Band", yaxis_title="Reps",
            plot_bgcolor=C["bg"], paper_bgcolor=C["bg"],
            font=dict(family="72,Helvetica Neue,Arial,sans-serif", size=12))
        fig_dist.update_yaxes(showgrid=True, gridcolor=C["grey"])
        st.plotly_chart(fig_dist, use_container_width=True)
    else:
        st.info("No quota-carrying reps in current filter.")


# ─────────────────────────────────────────────────────────────────────────────
# ROW 4 — REP PERFORMANCE TABLE
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='sh'>Rep Performance Detail</div>", unsafe_allow_html=True)

rep_show = rep_df.copy()
rep_show = rep_show.sort_values("Total_Credited", ascending=False)
_pct_disp = rep_show["Pct_to_Quota"].apply(fmt_pct)

rep_disp = pd.DataFrame({
    "Rep":              rep_show["Rep"],
    "Market":           rep_show["Market"],
    "Team":             rep_show["Team"],
    "Leader":           rep_show["Leader"],
    # Monetary columns converted to each rep's local currency for display.
    # Performance figures (CW_ARR etc.) come from SFDC "(converted)" fields
    # in USD; quota is also in USD.  fmt_local(usd_val, currency) applies
    # the 2026 budget FX rate and the correct local currency symbol.
    "CW ARR":           rep_show.apply(
        lambda r: fmt_local(r["CW_ARR"],           REP_CURRENCY.get(r["Rep"], "USD")), axis=1),
    "LTC":              rep_show.apply(
        lambda r: fmt_local(r["LTC_Credit"],        REP_CURRENCY.get(r["Rep"], "USD")), axis=1),
    "Retention":        rep_show.apply(
        lambda r: fmt_local(r["Retention_Credit"],  REP_CURRENCY.get(r["Rep"], "USD")), axis=1),
    "Complete/Ref":     rep_show.apply(
        lambda r: fmt_local(r["Complete_Credit"],   REP_CURRENCY.get(r["Rep"], "USD")), axis=1),
    "Total Credited":   rep_show.apply(
        lambda r: fmt_local(r["Total_Credited"],    REP_CURRENCY.get(r["Rep"], "USD")), axis=1),
    "Period Quota":     rep_show.apply(
        lambda r: fmt_local(r["Period_Quota"],      REP_CURRENCY.get(r["Rep"], "USD")), axis=1),
    "% to Quota":       _pct_disp,
    "CW Units":         rep_show["CW_Units"].astype(int),
})

def _style_row(row):
    """Colour-code % to Quota column."""
    styles = [""] * len(row)
    idx = rep_disp.columns.get_loc("% to Quota")
    pct_str = row["% to Quota"]
    try:
        pct = float(pct_str.replace("%","")) / 100
        if pct >= 1.0:     clr = "#E8F5E9"
        elif pct >= 0.75:  clr = "#FFF8E1"
        elif pct >= 0.50:  clr = "#FFF3E0"
        else:              clr = "#FFEBEE"
        styles[idx] = f"background-color:{clr}"
    except: pass
    return styles

styled = rep_disp.style.apply(_style_row, axis=1)
st.dataframe(styled, use_container_width=True, height=420, hide_index=True)


# ─────────────────────────────────────────────────────────────────────────────
# ROW 5 — MARKET DETAIL TABLE
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("<div class='sh'>Market Summary</div>", unsafe_allow_html=True)

mkt_disp = mkt_df.copy() if len(mkt_df) > 0 else pd.DataFrame(
    columns=["Market","CW_ARR","LTC_Credit","Retention_Credit",
             "Complete_Credit","Total_Credited","PL_Quota","Pct_to_PL","CW_Units"])

mkt_table = pd.DataFrame({
    "Market":           mkt_disp["Market"],
    "Currency":         "USD",   # all market rollups are in USD
    "CW ARR":           mkt_disp["CW_ARR"].round(0),
    "LTC":              mkt_disp["LTC_Credit"].round(0),
    "Retention":        mkt_disp["Retention_Credit"].round(0),
    "Complete/Ref":     mkt_disp["Complete_Credit"].round(0),
    "Total Credited":   mkt_disp["Total_Credited"].round(0),
    "PL Quota":         mkt_disp["PL_Quota"].round(0),
    "% to Plan":        mkt_disp["Pct_to_PL"].apply(fmt_pct),
    "CW Units":         mkt_disp["CW_Units"].astype(int) if len(mkt_disp) else pd.Series(dtype=int),
})
st.dataframe(mkt_table, use_container_width=True, height=180, hide_index=True)


# ─────────────────────────────────────────────────────────────────────────────
# ROW 6 — EXCEL DOWNLOAD
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div class='sh'>Export</div>", unsafe_allow_html=True)

@st.cache_data(show_spinner=False)
def build_excel(rep_df_json, ldr_df_json, mkt_df_json, period_lbl):
    # pd.read_json in pandas 3.x treats a bare string as a file path.
    # Wrap in io.StringIO so it is read as an in-memory JSON buffer.
    rep_df_e  = pd.read_json(io.StringIO(rep_df_json),  orient="records")
    ldr_df_e  = pd.read_json(io.StringIO(ldr_df_json),  orient="records")
    mkt_df_e  = pd.read_json(io.StringIO(mkt_df_json),  orient="records")

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        rep_df_e.to_excel(writer,  sheet_name="Rep Detail",    index=False)
        ldr_df_e.to_excel(writer,  sheet_name="Leader Summary",index=False)
        mkt_df_e.to_excel(writer,  sheet_name="Market Summary",index=False)

        # Pivot: rep x month (if month_cols available)
        meta_df = pd.DataFrame({"Metric": ["Period", "Generated"],
                                 "Value":  [period_lbl,
                                            datetime.now().strftime("%Y-%m-%d %H:%M")]})
        meta_df.to_excel(writer, sheet_name="Metadata", index=False)

    return buf.getvalue()

col_dl, col_debug = st.columns([2, 3])
with col_dl:
    try:
        xl_bytes = build_excel(
            rep_df.to_json(orient="records"),
            ldr_df.to_json(orient="records"),
            mkt_df.to_json(orient="records"),
            period_label)
        safe_period = re.sub(r"[^A-Za-z0-9_\-]", "_", period_label)
        mkt_slug = re.sub(r"\s+", "_", sel_market.replace("/","_"))
        fname = f"INT_SMB_{mkt_slug}_{safe_period}.xlsx"
        st.download_button(
            label="Download Excel Report",
            data=xl_bytes,
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    except Exception as e:
        st.error(f"Excel build error: {e}")

with col_debug:
    with st.expander("Diagnostics & Debug"):
        st.markdown(f"**Org row counts:** CW={org_d['Row_Counts']['CW']} | "
                    f"LTC={org_d['Row_Counts']['LTC']} | "
                    f"Retention={org_d['Row_Counts']['Retention']} | "
                    f"Complete={org_d['Row_Counts']['Complete']}")
        st.markdown(f"**Filtered reps:** {len(rep_df)} | "
                    f"**With quota:** {total_reps}")

        # Raw SFDC column names for CW
        if st.session_state.raw_cw is not None and len(st.session_state.raw_cw):
            st.markdown("**CW ARR columns:**")
            st.code(", ".join(st.session_state.raw_cw.columns.tolist()))

        # Report-level debug
        if "sf_post_debug" in st.session_state:
            for rid, msg in st.session_state.sf_post_debug.items():
                st.caption(f"{rid}: {msg}")

        # Currency note: non-standard local currencies are handled via
        # REP_CURRENCY + FX_RATES.  Chelsea Salonek (USD), Loreena Maguet
        # (EUR), Alan Donohoe (EUR), Lydia Holloway (AED) are all mapped.
        st.caption("Non-standard rep currencies: Chelsea Salonek (USD), "
                   "Loreena Maguet (EUR), Alan Donohoe (EUR), Lydia Holloway (AED)")
