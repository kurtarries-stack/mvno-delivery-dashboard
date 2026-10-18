import io
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="MVNO Delivery Control Room",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# DESIGN SYSTEM
# ============================================================

ORANGE = "#FF7A00"
ORANGE_LIGHT = "#FF9D3D"

BLACK = "#0B0B0B"
SIDEBAR = "#101010"
PANEL = "#171717"
PANEL_ALT = "#1D1D1D"
BORDER = "#2B2B2B"

WHITE = "#FFFFFF"
TEXT = "#F4F4F4"
MUTED = "#A8A8A8"

GREEN = "#22C55E"
AMBER = "#F59E0B"
RED = "#EF4444"
BLUE = "#3B82F6"
GREY = "#6B7280"


# ============================================================
# GLOBAL CSS
# ============================================================

st.markdown(
    f"""
    <style>

    .stApp {{
        background-color: {BLACK};
        color: {TEXT};
    }}

    [data-testid="stSidebar"] {{
        background-color: {SIDEBAR};
        border-right: 1px solid {BORDER};
    }}

    [data-testid="stSidebar"] * {{
        color: {TEXT};
    }}

    [data-testid="stMetric"] {{
        background-color: {PANEL};
        border: 1px solid {BORDER};
        border-left: 4px solid {ORANGE};
        padding: 18px;
        border-radius: 10px;
    }}

    [data-testid="stMetricLabel"] {{
        color: {MUTED};
        font-size: 0.85rem;
    }}

    [data-testid="stMetricValue"] {{
        color: {WHITE};
        font-weight: 700;
    }}

    .main-title {{
        font-size: 2.0rem;
        font-weight: 700;
        margin-bottom: 0;
        color: {WHITE};
    }}

    .main-subtitle {{
        color: {MUTED};
        margin-top: 0;
        margin-bottom: 25px;
    }}

    .orange {{
        color: {ORANGE};
    }}

    .section-heading {{
        color: {WHITE};
        font-size: 1.15rem;
        font-weight: 600;
        margin-top: 20px;
        margin-bottom: 10px;
    }}

    .small-note {{
        color: {MUTED};
        font-size: 0.82rem;
    }}

    .status-critical {{
        color: {RED};
        font-weight: 600;
    }}

    .status-warning {{
        color: {AMBER};
        font-weight: 600;
    }}

    .status-good {{
        color: {GREEN};
        font-weight: 600;
    }}

    div[data-testid="stSelectbox"] > div {{
        background-color: {PANEL};
    }}

    div[data-testid="stMultiSelect"] > div {{
        background-color: {PANEL};
    }}

    .stButton > button {{
        width: 100%;
        border: 1px solid {BORDER};
        background: {PANEL};
        color: {WHITE};
        border-radius: 7px;
    }}

    .stButton > button:hover {{
        border-color: {ORANGE};
        color: {ORANGE};
    }}

    .stDownloadButton > button {{
        border: 1px solid {ORANGE};
        background: {PANEL};
        color: {WHITE};
    }}

    hr {{
        border-color: {BORDER};
    }}

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return None

    value = str(value).strip()

    if value == "":
        return None

    return value


def normalise_column_name(value):
    if pd.isna(value):
        return ""

    return " ".join(str(value).strip().split())


def find_column(df, keywords):
    """
    Returns the first column whose name contains any supplied keyword.
    """

    for column in df.columns:
        name = str(column).lower().strip()

        for keyword in keywords:
            if keyword.lower() in name:
                return column

    return None


def status_is_complete(value):
    if pd.isna(value):
        return False

    status = str(value).strip().lower()

    completed = {
        "complete",
        "completed",
        "closed",
        "cancelled",
        "canceled",
    }

    return status in completed


# ============================================================
# LOAD EXECUTIVE MASTER DATASET
# ============================================================

@st.cache_data(show_spinner=False)
def load_executive(file_bytes):

    # Actual workbook structure:
    # Excel row 5 contains real headings
    # Excel row 6 contains secondary helper numbering/content
    # Data starts row 7

    raw = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name="Executive",
        header=None,
        engine="openpyxl",
        usecols="B:U",
    )

    # Excel row 5 = dataframe index 4
    headers = [
        normalise_column_name(x)
        for x in raw.iloc[4].tolist()
    ]

    df = raw.iloc[6:].copy()

    df.columns = headers

    # Remove duplicate/blank column names safely
    df = df.loc[:, df.columns != ""]

    # Power BI-friendly names
    rename_map = {
        "TASKS INFORMATION": "Deliverable",
        "DESCRIPTION": "Description",
        "FOCUS AREA": "Focus Area",
        "PRIORITY": "Priority Flag",
        "DELIVERY METHOD": "Delivery Method",
        "DELIVERY PHASE": "Delivery Phase",
        "OWNER": "Owner",
        "RESPONSIBLE": "Responsible",
        "REQUESTOR": "Requestor",
        "INITIATION DATE": "Initiation Date",
        "REVISED DUE DATE": "Revised Due Date",
        "ORIGINAL DUE DATE": "Original Due Date",
        "DAYS SINCE OPEN": "Days Since Open",
        "REMAINING DAYS": "Remaining Days",
        "CLIENT CLOSURE CONFIRMATION":
            "Client Closure Confirmation",
        "PHASE": "Phase",
        "STATUS": "Status",
        "PROGRESS (%)": "Progress",
        "MVNO": "MVNO",
        "SUMMARY REMARKS": "Summary Remarks",
    }

    df = df.rename(columns=rename_map)

    if "Deliverable" not in df.columns:
        raise ValueError(
            "The Executive sheet structure was not recognised."
        )

    # Remove blank rows
    df = df[
        df["Deliverable"].notna()
    ].copy()

    df["Deliverable"] = (
        df["Deliverable"]
        .astype(str)
        .str.strip()
    )

    df = df[
        df["Deliverable"] != ""
    ].copy()

    # Clean text columns
    text_columns = [
        "Deliverable",
        "Description",
        "Focus Area",
        "Priority Flag",
        "Delivery Method",
        "Delivery Phase",
        "Owner",
        "Responsible",
        "Requestor",
        "Client Closure Confirmation",
        "Phase",
        "Status",
        "MVNO",
        "Summary Remarks",
    ]

    for column in text_columns:
        if column in df.columns:
            df[column] = df[column].apply(clean_text)

    # Dates
    date_columns = [
        "Initiation Date",
        "Revised Due Date",
        "Original Due Date",
    ]

    for column in date_columns:
        if column in df.columns:
            df[column] = pd.to_datetime(
                df[column],
                errors="coerce",
            )

    # Progress
    if "Progress" in df.columns:
        df["Progress"] = pd.to_numeric(
            df["Progress"],
            errors="coerce",
        )

    # Effective due date
    df["Effective Due Date"] = (
        df["Revised Due Date"]
        .fillna(df["Original Due Date"])
    )

    # Completion
    df["Is Complete"] = (
        df["Status"]
        .apply(status_is_complete)
    )

    today = pd.Timestamp.today().normalize()

    # Overdue
    df["Is Overdue"] = (
        df["Effective Due Date"].notna()
        & (df["Effective Due Date"] < today)
        & (~df["Is Complete"])
    )

    # Due next seven days
    df["Due Next 7 Days"] = (
        df["Effective Due Date"].notna()
        & (df["Effective Due Date"] >= today)
        & (
            df["Effective Due Date"]
            <= today + pd.Timedelta(days=7)
        )
        & (~df["Is Complete"])
    )

    # Due this week
    monday = (
        today
        - pd.Timedelta(days=today.weekday())
    )

    sunday = monday + pd.Timedelta(days=6)

    df["Due This Week"] = (
        df["Effective Due Date"].notna()
        & (df["Effective Due Date"] >= monday)
        & (df["Effective Due Date"] <= sunday)
        & (~df["Is Complete"])
    )

    # Progress bands
    def progress_band(value):

        if pd.isna(value):
            return "Not captured"

        if value >= 1:
            return "100%"

        if value >= 0.75:
            return "75–99%"

        if value >= 0.50:
            return "50–74%"

        if value >= 0.25:
            return "25–49%"

        if value > 0:
            return "1–24%"

        return "0%"

    df["Progress Band"] = (
        df["Progress"]
        .apply(progress_band)
    )

    # Due bucket
    def due_bucket(row):

        if row["Is Complete"]:
            return "Complete"

        if pd.isna(row["Effective Due Date"]):
            return "No due date"

        if row["Is Overdue"]:
            return "Overdue"

        if row["Due Next 7 Days"]:
            return "Due next 7 days"

        return "Future"

    df["Due Bucket"] = df.apply(
        due_bucket,
        axis=1,
    )

    df["Deliverable ID"] = np.arange(
        1,
        len(df) + 1,
    )

    return df.reset_index(drop=True)


# ============================================================
# BILLING
# ============================================================

@st.cache_data(show_spinner=False)
def load_billing(file_bytes):

    raw = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name="Billing",
        header=None,
        engine="openpyxl",
        usecols="B:J",
    )

    # Billing row 8 contains headers
    headers = [
        normalise_column_name(x)
        for x in raw.iloc[7].tolist()
    ]

    df = raw.iloc[8:].copy()
    df.columns = headers

    if "Service" not in df.columns:
        return pd.DataFrame()

    df = df[
        df["Service"].notna()
    ].copy()

    text_columns = [
        "Service",
        "MVNO",
        "Customer",
        "Validation Status",
        "Comments",
    ]

    for column in text_columns:
        if column in df.columns:
            df[column] = df[column].apply(clean_text)

    numeric = [
        "Customers Affected",
        "Unbilled Records",
        "Rand Value",
    ]

    for column in numeric:
        if column in df.columns:
            df[column] = pd.to_numeric(
                df[column],
                errors="coerce",
            )

    if "Date Logged" in df.columns:
        df["Date Logged"] = pd.to_datetime(
            df["Date Logged"],
            errors="coerce",
        )

    return df.reset_index(drop=True)


# ============================================================
# QUERY WORKBOOK SHEET
# ============================================================

@st.cache_data(show_spinner=False)
def load_generic_sheet(
    file_bytes,
    sheet,
):

    try:

        return pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet,
            engine="openpyxl",
        )

    except Exception:

        return pd.DataFrame()


# ============================================================
# FILTERS
# ============================================================

def apply_global_filters(
    df,
    key_prefix,
):

    filtered = df.copy()

    st.sidebar.markdown("---")
    st.sidebar.markdown("### Filters")

    search = st.sidebar.text_input(
        "Search",
        placeholder="Search deliverables...",
        key=f"{key_prefix}_search",
    )

    if search:

        mask = (
            filtered.astype(str)
            .apply(
                lambda column:
                column.str.contains(
                    search,
                    case=False,
                    na=False,
                )
            )
            .any(axis=1)
        )

        filtered = filtered[mask]

    filter_columns = [
        "MVNO",
        "Owner",
        "Responsible",
        "Status",
        "Focus Area",
        "Priority Flag",
        "Delivery Phase",
    ]

    for column in filter_columns:

        if column not in filtered.columns:
            continue

        available = sorted(
            filtered[column]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        if len(available) == 0:
            continue

        selection = st.sidebar.multiselect(
            column,
            available,
            key=f"{key_prefix}_{column}",
        )

        if selection:

            filtered = filtered[
                filtered[column]
                .astype(str)
                .isin(selection)
            ]

    return filtered


# ============================================================
# PLOTLY THEME
# ============================================================

def style_chart(
    fig,
    height=360,
):

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor=PANEL,
        plot_bgcolor=PANEL,
        font_color=TEXT,
        height=height,
        margin=dict(
            l=20,
            r=20,
            t=50,
            b=20,
        ),
        legend=dict(
            bgcolor="rgba(0,0,0,0)"
        ),
    )

    return fig


# ============================================================
# KPI CARDS
# ============================================================

def render_kpis(df):

    total = len(df)

    open_items = int(
        (~df["Is Complete"]).sum()
    )

    overdue = int(
        df["Is Overdue"].sum()
    )

    critical = int(
        df["Status"]
        .fillna("")
        .astype(str)
        .str.contains(
            "Critical|Delayed",
            case=False,
            regex=True,
        )
        .sum()
    )

    completion_rate = (
        df["Is Complete"].mean()
        if len(df)
        else 0
    )

    avg_progress = (
        df["Progress"].mean()
        if df["Progress"].notna().any()
        else 0
    )

    cols = st.columns(6)

    cols[0].metric(
        "Deliverables",
        f"{total:,}",
    )

    cols[1].metric(
        "Open",
        f"{open_items:,}",
    )

    cols[2].metric(
        "Overdue",
        f"{overdue:,}",
    )

    cols[3].metric(
        "Critical / Delayed",
        f"{critical:,}",
    )

    cols[4].metric(
        "Completion",
        f"{completion_rate:.1%}",
    )

    cols[5].metric(
        "Average Progress",
        f"{avg_progress:.1%}",
    )


# ============================================================
# EXECUTIVE PAGE
# ============================================================

def executive_page(df):

    st.markdown(
        """
        <div class="main-title">
        MVNO <span class="orange">Delivery Control Room</span>
        </div>

        <div class="main-subtitle">
        Executive portfolio overview
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_kpis(df)

    st.markdown("")

    left, right = st.columns(2)

    # Status distribution
    with left:

        st.markdown(
            '<div class="section-heading">Status Distribution</div>',
            unsafe_allow_html=True,
        )

        status = (
            df["Status"]
            .fillna("Not captured")
            .value_counts()
            .reset_index()
        )

        status.columns = [
            "Status",
            "Deliverables",
        ]

        fig = px.bar(
            status,
            x="Deliverables",
            y="Status",
            orientation="h",
            color="Status",
            color_discrete_sequence=[
                ORANGE,
                ORANGE_LIGHT,
                RED,
                AMBER,
                GREEN,
                BLUE,
                GREY,
            ],
        )

        fig.update_layout(
            showlegend=False
        )

        style_chart(fig)

        st.plotly_chart(
            fig,
            use_container_width=True,
        )

    # MVNO workload
    with right:

        st.markdown(
            '<div class="section-heading">Portfolio by MVNO</div>',
            unsafe_allow_html=True,
        )

        mvno = (
            df["MVNO"]
            .fillna("Not captured")
            .value_counts()
            .reset_index()
        )

        mvno.columns = [
            "MVNO",
            "Deliverables",
        ]

        fig = px.bar(
            mvno,
            x="Deliverables",
            y="MVNO",
            orientation="h",
            color_discrete_sequence=[
                ORANGE
            ],
        )

        fig.update_layout(
            showlegend=False
        )

        style_chart(fig)

        st.plotly_chart(
            fig,
            use_container_width=True,
        )

    left, right = st.columns(2)

    # Progress
    with left:

        st.markdown(
            '<div class="section-heading">Progress Distribution</div>',
            unsafe_allow_html=True,
        )

        order = [
            "0%",
            "1–24%",
            "25–49%",
            "50–74%",
            "75–99%",
            "100%",
            "Not captured",
        ]

        progress = (
            df["Progress Band"]
            .value_counts()
            .reindex(
                order,
                fill_value=0,
            )
            .reset_index()
        )

        progress.columns = [
            "Progress Band",
            "Deliverables",
        ]

        fig = px.bar(
            progress,
            x="Progress Band",
            y="Deliverables",
            color_discrete_sequence=[
                ORANGE
            ],
        )

        style_chart(fig)

        st.plotly_chart(
            fig,
            use_container_width=True,
        )

    # Due position
    with right:

        st.markdown(
            '<div class="section-heading">Due Position</div>',
            unsafe_allow_html=True,
        )

        due = (
            df["Due Bucket"]
            .value_counts()
            .reset_index()
        )

        due.columns = [
            "Due Bucket",
            "Deliverables",
        ]

        fig = px.pie(
            due,
            values="Deliverables",
            names="Due Bucket",
            hole=0.65,
            color_discrete_sequence=[
                RED,
                AMBER,
                ORANGE,
                GREEN,
                GREY,
            ],
        )

        style_chart(fig)

        st.plotly_chart(
            fig,
            use_container_width=True,
        )

    watchlist(
        df,
        "Executive Risk Watchlist",
    )


# ============================================================
# WATCHLIST
# ============================================================

def watchlist(
    df,
    title="Watchlist",
):

    st.markdown(
        f'<div class="section-heading">{title}</div>',
        unsafe_allow_html=True,
    )

    high_risk = df[
        df["Is Overdue"]
        |
        df["Status"]
        .fillna("")
        .astype(str)
        .str.contains(
            "Critical|Delayed",
            case=False,
            regex=True,
        )
    ].copy()

    show_columns = [
        "MVNO",
        "Deliverable",
        "Focus Area",
        "Responsible",
        "Status",
        "Progress",
        "Effective Due Date",
        "Summary Remarks",
    ]

    show_columns = [
        c
        for c in show_columns
        if c in high_risk.columns
    ]

    result = high_risk[
        show_columns
    ].copy()

    if "Progress" in result.columns:
        result["Progress"] = (
            result["Progress"]
            .apply(
                lambda x:
                f"{x:.0%}"
                if pd.notna(x)
                else ""
            )
        )

    if "Effective Due Date" in result.columns:
        result[
            "Effective Due Date"
        ] = (
            result[
                "Effective Due Date"
            ]
            .dt.strftime("%Y-%m-%d")
        )

    st.dataframe(
        result,
        use_container_width=True,
        height=420,
        hide_index=True,
    )


# ============================================================
# WORKBENCH PAGE
# ============================================================

def workbench_page(df):

    st.markdown(
        """
        <div class="main-title">
        <span class="orange">Workbench</span>
        </div>

        <div class="main-subtitle">
        Operational triage and immediate delivery focus
        </div>
        """,
        unsafe_allow_html=True,
    )

    open_df = df[
        ~df["Is Complete"]
    ].copy()

    overdue = open_df[
        open_df["Is Overdue"]
    ]

    next7 = open_df[
        open_df["Due Next 7 Days"]
    ]

    zero_progress = open_df[
        open_df["Progress"]
        .fillna(0)
        .eq(0)
    ]

    no_due_date = open_df[
        open_df["Effective Due Date"]
        .isna()
    ]

    a, b, c, d, e = st.columns(5)

    a.metric(
        "Open",
        len(open_df),
    )

    b.metric(
        "Overdue",
        len(overdue),
    )

    c.metric(
        "Due Next 7 Days",
        len(next7),
    )

    d.metric(
        "Zero Progress",
        len(zero_progress),
    )

    e.metric(
        "No Due Date",
        len(no_due_date),
    )

    st.markdown("")

    view = st.radio(
        "Workbench view",
        [
            "Focus Now",
            "All Open",
            "Overdue",
            "Due Next 7 Days",
            "Zero Progress",
            "No Due Date",
        ],
        horizontal=True,
    )

    if view == "All Open":
        work = open_df

    elif view == "Overdue":
        work = overdue

    elif view == "Due Next 7 Days":
        work = next7

    elif view == "Zero Progress":
        work = zero_progress

    elif view == "No Due Date":
        work = no_due_date

    else:

        work = open_df[
            open_df["Is Overdue"]
            |
            open_df["Due Next 7 Days"]
            |
            open_df["Status"]
            .fillna("")
            .astype(str)
            .str.contains(
                "Critical|Delayed",
                case=False,
                regex=True,
            )
        ]

    detail_table(
        work,
        title=view,
    )


# ============================================================
# PRIORITY VIEW
# ============================================================

def priority_page(df):

    st.markdown(
        """
        <div class="main-title">
        <span class="orange">Priority View</span>
        </div>

        <div class="main-subtitle">
        Risk, urgency and delivery attention
        </div>
        """,
        unsafe_allow_html=True,
    )

    open_df = df[
        ~df["Is Complete"]
    ].copy()

    # Dashboard-native attention score.
    #
    # This does NOT attempt to overwrite the workbook's
    # original priority engine. It creates a transparent,
    # simple operational prioritisation for the web dashboard.

    open_df["Attention Score"] = 0

    open_df.loc[
        open_df["Is Overdue"],
        "Attention Score",
    ] += 40

    open_df.loc[
        open_df["Due Next 7 Days"],
        "Attention Score",
    ] += 20

    open_df.loc[
        open_df["Status"]
        .fillna("")
        .str.contains(
            "Critical",
            case=False,
        ),
        "Attention Score",
    ] += 40

    open_df.loc[
        open_df["Status"]
        .fillna("")
        .str.contains(
            "Delayed",
            case=False,
        ),
        "Attention Score",
    ] += 30

    open_df.loc[
        open_df["Progress"]
        .fillna(0)
        .eq(0),
        "Attention Score",
    ] += 10

    open_df.loc[
        open_df["Effective Due Date"]
        .isna(),
        "Attention Score",
    ] += 10

    open_df["Attention Level"] = pd.cut(
        open_df["Attention Score"],
        bins=[
            -1,
            19,
            39,
            59,
            999,
        ],
        labels=[
            "Monitor",
            "Medium",
            "High",
            "Immediate",
        ],
    )

    open_df = open_df.sort_values(
        [
            "Attention Score",
            "Effective Due Date",
        ],
        ascending=[
            False,
            True,
        ],
        na_position="last",
    )

    a, b, c, d = st.columns(4)

    a.metric(
        "Immediate",
        int(
            (
                open_df["Attention Level"]
                == "Immediate"
            ).sum()
        ),
    )

    b.metric(
        "High",
        int(
            (
                open_df["Attention Level"]
                == "High"
            ).sum()
        ),
    )

    c.metric(
        "Medium",
        int(
            (
                open_df["Attention Level"]
                == "Medium"
            ).sum()
        ),
    )

    d.metric(
        "Monitor",
        int(
            (
                open_df["Attention Level"]
                == "Monitor"
            ).sum()
        ),
    )

    detail_table(
        open_df,
        title="Prioritised Deliverables",
        extra_columns=[
            "Attention Score",
            "Attention Level",
        ],
    )


# ============================================================
# WEEKLY VIEW
# ============================================================

def weekly_page(df):

    st.markdown(
        """
        <div class="main-title">
        <span class="orange">Weekly View</span>
        </div>

        <div class="main-subtitle">
        Current delivery window
        </div>
        """,
        unsafe_allow_html=True,
    )

    today = pd.Timestamp.today().normalize()

    week_start = (
        today
        - pd.Timedelta(days=today.weekday())
    )

    week_end = (
        week_start
        + pd.Timedelta(days=6)
    )

    next_week_end = (
        week_end
        + pd.Timedelta(days=7)
    )

    overdue = df[
        df["Is Overdue"]
    ]

    due_this_week = df[
        (
            df["Effective Due Date"]
            >= week_start
        )
        &
        (
            df["Effective Due Date"]
            <= week_end
        )
        &
        (~df["Is Complete"])
    ]

    due_next_week = df[
        (
            df["Effective Due Date"]
            > week_end
        )
        &
        (
            df["Effective Due Date"]
            <= next_week_end
        )
        &
        (~df["Is Complete"])
    ]

    a, b, c = st.columns(3)

    a.metric(
        "Overdue",
        len(overdue),
    )

    b.metric(
        "Due This Week",
        len(due_this_week),
    )

    c.metric(
        "Due Next Week",
        len(due_next_week),
    )

    tab1, tab2, tab3 = st.tabs(
        [
            "Overdue",
            "This Week",
            "Next Week",
        ]
    )

    with tab1:
        detail_table(
            overdue,
            "Overdue Deliverables",
        )

    with tab2:
        detail_table(
            due_this_week,
            "Due This Week",
        )

    with tab3:
        detail_table(
            due_next_week,
            "Due Next Week",
        )


# ============================================================
# TEAM DASHBOARD
# ============================================================

def team_page(df):

    st.markdown(
        """
        <div class="main-title">
        Team <span class="orange">Dashboard</span>
        </div>

        <div class="main-subtitle">
        Delivery workload and performance
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_kpis(df)

    left, right = st.columns(2)

    with left:

        workload = (
            df["Responsible"]
            .fillna("Not assigned")
            .value_counts()
            .head(15)
            .reset_index()
        )

        workload.columns = [
            "Responsible",
            "Deliverables",
        ]

        fig = px.bar(
            workload,
            x="Deliverables",
            y="Responsible",
            orientation="h",
            color_discrete_sequence=[
                ORANGE
            ],
            title="Workload by Responsible",
        )

        style_chart(fig, 450)

        st.plotly_chart(
            fig,
            use_container_width=True,
        )

    with right:

        overdue_team = (
            df[
                df["Is Overdue"]
            ]["Responsible"]
            .fillna("Not assigned")
            .value_counts()
            .head(15)
            .reset_index()
        )

        overdue_team.columns = [
            "Responsible",
            "Overdue",
        ]

        fig = px.bar(
            overdue_team,
            x="Overdue",
            y="Responsible",
            orientation="h",
            color_discrete_sequence=[
                RED
            ],
            title="Overdue Pressure",
        )

        style_chart(fig, 450)

        st.plotly_chart(
            fig,
            use_container_width=True,
        )


# ============================================================
# CLIENT / PORTFOLIO PAGE
# ============================================================

def portfolio_page(
    df,
    title,
):

    st.markdown(
        f"""
        <div class="main-title">
        <span class="orange">{title}</span>
        </div>

        <div class="main-subtitle">
        Delivery portfolio
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_kpis(df)

    left, right = st.columns(2)

    with left:

        status = (
            df["Status"]
            .fillna("Not captured")
            .value_counts()
            .reset_index()
        )

        status.columns = [
            "Status",
            "Deliverables",
        ]

        fig = px.bar(
            status,
            x="Deliverables",
            y="Status",
            orientation="h",
            color_discrete_sequence=[
                ORANGE
            ],
            title="Status",
        )

        style_chart(fig)

        st.plotly_chart(
            fig,
            use_container_width=True,
        )

    with right:

        focus = (
            df["Focus Area"]
            .fillna("Not captured")
            .value_counts()
            .head(10)
            .reset_index()
        )

        focus.columns = [
            "Focus Area",
            "Deliverables",
        ]

        fig = px.bar(
            focus,
            x="Deliverables",
            y="Focus Area",
            orientation="h",
            color_discrete_sequence=[
                ORANGE_LIGHT
            ],
            title="Focus Areas",
        )

        style_chart(fig)

        st.plotly_chart(
            fig,
            use_container_width=True,
        )

    watchlist(
        df,
        f"{title} Watchlist",
    )

    detail_table(
        df,
        f"{title} Deliverables",
    )


# ============================================================
# BILLING PAGE
# ============================================================

def billing_page(billing):

    st.markdown(
        """
        <div class="main-title">
        <span class="orange">Billing</span>
        </div>

        <div class="main-subtitle">
        Unbilled records and financial exposure
        </div>
        """,
        unsafe_allow_html=True,
    )

    if billing.empty:

        st.info(
            "No billing register records were found."
        )

        return

    total_records = (
        billing["Unbilled Records"].sum()
        if "Unbilled Records" in billing
        else 0
    )

    rand_value = (
        billing["Rand Value"].sum()
        if "Rand Value" in billing
        else 0
    )

    customers = (
        billing["Customers Affected"].sum()
        if "Customers Affected" in billing
        else 0
    )

    a, b, c = st.columns(3)

    a.metric(
        "Customers Affected",
        f"{customers:,.0f}",
    )

    b.metric(
        "Unbilled Records",
        f"{total_records:,.0f}",
    )

    c.metric(
        "Rand Exposure",
        f"R {rand_value:,.2f}",
    )

    st.markdown("")

    st.dataframe(
        billing,
        use_container_width=True,
        hide_index=True,
        height=520,
    )


# ============================================================
# QUALITY PAGE
# ============================================================

def quality_page(df):

    st.markdown(
        """
        <div class="main-title">
        Data <span class="orange">Quality</span>
        </div>

        <div class="main-subtitle">
        Tracker completeness and data-quality controls
        </div>
        """,
        unsafe_allow_html=True,
    )

    checks = {
        "Missing Initiation Date":
            df["Initiation Date"].isna(),

        "Missing Due Date":
            df["Effective Due Date"].isna(),

        "Missing Responsible":
            df["Responsible"].isna(),

        "Missing Status":
            df["Status"].isna(),

        "Missing Progress":
            df["Progress"].isna(),

        "Open with 100% Progress":
            (
                (~df["Is Complete"])
                &
                (df["Progress"] >= 1)
            ),

        "Complete Below 100%":
            (
                df["Is Complete"]
                &
                (df["Progress"] < 1)
            ),
    }

    rows = []

    for check, mask in checks.items():

        rows.append(
            {
                "Check": check,
                "Issues": int(mask.sum()),
            }
        )

    quality = pd.DataFrame(rows)

    total_issues = int(
        quality["Issues"].sum()
    )

    impacted_rows = pd.Series(
        False,
        index=df.index,
    )

    for mask in checks.values():
        impacted_rows = (
            impacted_rows | mask
        )

    clean_rows = (
        len(df)
        - int(impacted_rows.sum())
    )

    clean_rate = (
        clean_rows / len(df)
        if len(df)
        else 0
    )

    a, b, c = st.columns(3)

    a.metric(
        "Quality Issues",
        f"{total_issues:,}",
    )

    b.metric(
        "Clean Rows",
        f"{clean_rows:,}",
    )

    c.metric(
        "Clean Rate",
        f"{clean_rate:.1%}",
    )

    fig = px.bar(
        quality.sort_values(
            "Issues",
            ascending=True,
        ),
        x="Issues",
        y="Check",
        orientation="h",
        color_discrete_sequence=[
            ORANGE
        ],
    )

    style_chart(fig, 420)

    st.plotly_chart(
        fig,
        use_container_width=True,
    )


# ============================================================
# DETAIL TABLE
# ============================================================

def detail_table(
    df,
    title="Detail",
    extra_columns=None,
):

    st.markdown(
        f'<div class="section-heading">{title}</div>',
        unsafe_allow_html=True,
    )

    columns = [
        "MVNO",
        "Deliverable",
        "Focus Area",
        "Priority Flag",
        "Delivery Method",
        "Delivery Phase",
        "Owner",
        "Responsible",
        "Requestor",
        "Status",
        "Progress",
        "Initiation Date",
        "Effective Due Date",
        "Days Since Open",
        "Remaining Days",
        "Summary Remarks",
    ]

    if extra_columns:

        columns = (
            extra_columns
            + columns
        )

    columns = [
        c
        for c in columns
        if c in df.columns
    ]

    display = (
        df[columns]
        .copy()
    )

    if "Progress" in display.columns:

        display["Progress"] = (
            display["Progress"]
            .apply(
                lambda x:
                f"{x:.0%}"
                if pd.notna(x)
                else ""
            )
        )

    date_columns = [
        "Initiation Date",
        "Effective Due Date",
    ]

    for column in date_columns:

        if column in display.columns:

            display[column] = (
                pd.to_datetime(
                    display[column],
                    errors="coerce",
                )
                .dt.strftime(
                    "%Y-%m-%d"
                )
            )

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        height=540,
    )

    csv = display.to_csv(
        index=False
    ).encode("utf-8-sig")

    st.download_button(
        label="Download filtered view",
        data=csv,
        file_name=(
            title
            .lower()
            .replace(" ", "_")
            + ".csv"
        ),
        mime="text/csv",
    )


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.markdown(
    """
    ## 🟧 MVNO
    ### Delivery Control Room
    """
)

st.sidebar.caption(
    "Delivery portfolio intelligence"
)

st.sidebar.markdown("---")


# ============================================================
# WORKBOOK UPLOAD
# ============================================================

uploaded_file = st.sidebar.file_uploader(
    "Upload MVNO workbook",
    type=[
        "xlsx",
        "xlsm",
    ],
)

if uploaded_file is None:

    st.markdown(
        """
        <div class="main-title">
        MVNO <span class="orange">Delivery Control Room</span>
        </div>

        <div class="main-subtitle">
        Upload the MVNO Delivery workbook to begin.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.info(
        "Use **Upload MVNO workbook** in the left panel."
    )

    st.markdown(
        """
        ### What's inside

        Once the workbook is uploaded, the dashboard will build:

        - Executive portfolio overview
        - Operational Workbench
        - Weekly delivery view
        - Team workload dashboard
        - Priority view
        - Roadmap and Internal portfolios
        - Individual MVNO views
        - Billing exposure
        - Data-quality controls
        """
    )

    st.stop()


file_bytes = uploaded_file.getvalue()


# ============================================================
# LOAD DATA
# ============================================================

try:

    with st.spinner(
        "Reading the MVNO delivery portfolio..."
    ):

        deliverables = load_executive(
            file_bytes
        )

        billing = load_billing(
            file_bytes
        )

except Exception as error:

    st.error(
        "The workbook could not be interpreted."
    )

    st.exception(error)

    st.stop()


# ============================================================
# NAVIGATION
# ============================================================

st.sidebar.success(
    f"{len(deliverables):,} deliverables loaded"
)

navigation = [
    "Executive",
    "Workbench",
    "Weekly View",
    "Team Dashboard",
    "Priority View",

    "── PORTFOLIO ──",

    "Roadmap",
    "Internal",

    "── MVNOs ──",

    "Capitec",
    "FNB",
    "C-Connect",
    "FREI",
    "Melon",
    "Asuer",
    "MVN-X",
    "DevTel",
    "Sakeng",
    "Knect",
    "OMC",

    "── COMMERCIAL ──",

    "Billing",

    "── GOVERNANCE ──",

    "Quality",
]

page = st.sidebar.radio(
    "Navigation",
    navigation,
    label_visibility="collapsed",
)

# Prevent separators becoming pages
if page.startswith("──"):

    page = "Executive"


# ============================================================
# FILTER DATA FOR SPECIFIC PORTFOLIO PAGE
# ============================================================

filtered = deliverables.copy()

client_pages = [
    "Capitec",
    "FNB",
    "C-Connect",
    "FREI",
    "Melon",
    "Asuer",
    "MVN-X",
    "DevTel",
    "Sakeng",
    "Knect",
    "OMC",
]

if page in client_pages:

    filtered = filtered[
        filtered["MVNO"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
        ==
        page.casefold()
    ].copy()


# ============================================================
# ROADMAP / INTERNAL
# ============================================================

if page == "Roadmap":

    filtered = filtered[
        filtered["Delivery Method"]
        .fillna("")
        .astype(str)
        .str.contains(
            "Roadmap",
            case=False,
            na=False,
        )
    ].copy()


elif page == "Internal":

    filtered = filtered[
        filtered["MVNO"]
        .fillna("")
        .astype(str)
        .str.contains(
            "Internal",
            case=False,
            na=False,
        )
        |
        filtered["Delivery Method"]
        .fillna("")
        .astype(str)
        .str.contains(
            "Internal",
            case=False,
            na=False,
        )
    ].copy()


# ============================================================
# GLOBAL FILTERS
# ============================================================

# Billing uses its own data model
if page not in [
    "Billing",
    "Quality",
]:

    filtered = apply_global_filters(
        filtered,
        page.replace(
            " ",
            "_",
        ),
    )


# ============================================================
# PAGE ROUTER
# ============================================================

if page == "Executive":

    executive_page(filtered)


elif page == "Workbench":

    workbench_page(filtered)


elif page == "Weekly View":

    weekly_page(filtered)


elif page == "Team Dashboard":

    team_page(filtered)


elif page == "Priority View":

    priority_page(filtered)


elif page == "Roadmap":

    portfolio_page(
        filtered,
        "Roadmap",
    )


elif page == "Internal":

    portfolio_page(
        filtered,
        "Internal",
    )


elif page in client_pages:

    portfolio_page(
        filtered,
        page,
    )


elif page == "Billing":

    billing_page(billing)


elif page == "Quality":

    quality_page(
        deliverables
    )


# ============================================================
# FOOTER
# ============================================================

st.sidebar.markdown("---")

st.sidebar.caption(
    "MVNO Delivery Control Room"
)

st.sidebar.caption(
    "Streamlit • GitHub"
)
