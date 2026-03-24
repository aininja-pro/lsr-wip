"""
CSS Theme for LSR_Pay Application

Injects custom CSS to give the Streamlit app a professional appearance
using LSR Multifamily brand colors.
"""

import streamlit as st

# Brand colors
LSR_BLUE = "#1E3A5F"
LSR_RED = "#C41E3A"


def inject_custom_css():
    """Inject custom CSS theme. Call once at the top of the app."""
    st.markdown("""
    <style>
        /* Background */
        .stApp {
            background-color: #FAFAFA;
        }

        /* Headings */
        h1, h2, h3, h4 {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            color: #1E3A5F;
        }

        h1 {
            font-size: 1.8rem !important;
            font-weight: 700 !important;
        }

        /* Body text */
        p, li, span, label, .stMarkdown {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            color: #1A1A1A;
        }

        /* Secondary text */
        .stCaption, [data-testid="stCaptionContainer"] {
            color: #6B7280;
        }

        /* Primary buttons */
        .stButton > button[kind="primary"],
        .stButton > button[data-testid="stBaseButton-primary"] {
            background-color: #1E3A5F !important;
            color: white !important;
            border: none;
            border-radius: 6px;
            font-weight: 600;
            padding: 0.5rem 1.5rem;
        }
        .stButton > button[kind="primary"]:hover,
        .stButton > button[data-testid="stBaseButton-primary"]:hover {
            background-color: #2a4f7a !important;
            color: white !important;
        }
        .stButton > button[kind="primary"] p,
        .stButton > button[data-testid="stBaseButton-primary"] p {
            color: white !important;
        }

        /* Download buttons */
        .stDownloadButton > button {
            background-color: #1E3A5F;
            color: white;
            border: none;
            border-radius: 6px;
            font-weight: 600;
            padding: 0.5rem 1.5rem;
        }
        .stDownloadButton > button:hover {
            background-color: #2a4f7a;
            color: white;
        }

        /* File uploader */
        [data-testid="stFileUploader"] {
            border: 2px dashed #E5E7EB;
            border-radius: 8px;
            padding: 1rem;
        }
        [data-testid="stFileUploader"]:hover {
            border-color: #1E3A5F;
        }

        /* DataFrames */
        [data-testid="stDataFrame"] {
            border: 1px solid #E5E7EB;
            border-radius: 8px;
        }

        /* Tabs styling */
        .stTabs [data-baseweb="tab-list"] {
            gap: 2rem;
            border-bottom: 2px solid #E5E7EB;
        }
        .stTabs [data-baseweb="tab"] {
            font-weight: 600;
            font-size: 1rem;
            padding: 0.75rem 0;
        }
        .stTabs [aria-selected="true"] {
            color: #1E3A5F;
        }

        /* Card-like containers */
        [data-testid="stExpander"] {
            background-color: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }

        /* Sidebar */
        [data-testid="stSidebar"] {
            background-color: #FFFFFF;
            border-right: 1px solid #E5E7EB;
        }

        /* Success/warning/error messages */
        .stAlert [data-testid="stAlertContentWarning"] {
            border-left-color: #C41E3A;
        }

        /* Metric cards */
        [data-testid="stMetric"] {
            background-color: #FFFFFF;
            border: 1px solid #E5E7EB;
            border-radius: 8px;
            padding: 1rem;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }

        /* Horizontal rule */
        hr {
            border-color: #E5E7EB;
        }
    </style>
    """, unsafe_allow_html=True)
