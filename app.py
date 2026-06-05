"""
SmartClaim AI — Streamlit Dashboard
A premium dark-themed UI for Arabic vehicle accident cost prediction.
"""

import streamlit as st
import json
import os
import sys
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# ---------------------------------------------------------------------------
# Ensure working directory is project root so relative paths in src/ work
# ---------------------------------------------------------------------------
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# ---------------------------------------------------------------------------
# Page configuration (must be first Streamlit call)
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="SmartClaim AI",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Load custom CSS
# ---------------------------------------------------------------------------
def load_css():
    css_path = os.path.join(os.path.dirname(__file__), "style.css")
    if os.path.exists(css_path):
        with open(css_path) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css()

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
TIER_CONFIG = {
    "Very Light":                {"tier": 1, "range": "0 – 1,000",       "color": "#4ade80"},
    "Light (Scratches)":         {"tier": 2, "range": "1,000 – 3,000",   "color": "#a3e635"},
    "Minor":                     {"tier": 3, "range": "3,000 – 6,000",   "color": "#facc15"},
    "Moderate":                  {"tier": 4, "range": "4,000 – 8,000",   "color": "#fb923c"},
    "Moderate (Multiple Parts)": {"tier": 5, "range": "8,000 – 15,000",  "color": "#f97316"},
    "Severe":                    {"tier": 6, "range": "15,000 – 25,000", "color": "#ef4444"},
    "Severe (Critical/Structural)": {"tier": 7, "range": "25,000 – 60,000+", "color": "#dc2626"},
}

EXAMPLE_CASES = [
    ("Very Light",  "صدمة خفيفة جداً بدون أضرار واضحة"),
    ("Scratches",   "خدوش بسيطة في الرفرف الخلفي"),
    ("Moderate",    "صدم من الخلف تضرر الصدام والشمعة يحتاج إصلاح"),
    ("Severe",      "تدمير قوي في الواجهة الأمامية والهيكل والصدام والكبوت"),
    ("Critical",    "حادث شديد، تضررت الشاصيه والماكينة بالكامل مع تلف الصدام"),
]

# ---------------------------------------------------------------------------
# Cached loaders
# ---------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def load_ml_pipeline():
    """Pre-loads the ML models into memory once at startup."""
    try:
        from src.predict import _get_model, _get_embedder
        _ = _get_model()  # Loads best_model.joblib
        _ = _get_embedder().model  # Loads SentenceTransformer (paraphrase-multilingual-mpnet)
        return True
    except Exception:
        return False

# Trigger preload in the background
_ = load_ml_pipeline()

@st.cache_data(show_spinner=False)
def load_metrics():
    path = "models/metrics.json"
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return None

@st.cache_data(show_spinner=False)
def load_dataset():
    try:
        from src.preprocess import load_data
        return load_data()
    except Exception:
        return None

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def hero(title: str, subtitle: str):
    st.markdown(
        f"""<div class="hero-banner">
            <h1>{title}</h1>
            <p>{subtitle}</p>
        </div>""",
        unsafe_allow_html=True,
    )

def kpi_cards(metrics: dict):
    items = [
        ("📉", f"{metrics['MAE']:,.0f}",  "Mean Absolute Error (SAR)", "#00d4aa"),
        ("📈", f"{metrics['R2']:.3f}",     "R² Score",                   "#667eea"),
        ("📊", f"{metrics['RMSE']:,.0f}",  "RMSE (SAR)",                 "#f97316"),
        ("📐", f"{metrics['MAPE']*100:.1f}%", "MAPE",                    "#ef4444"),
    ]
    cards = ""
    for icon, value, label, color in items:
        cards += f"""
        <div class="kpi-card" style="--kpi-color:{color}">
            <div class="kpi-icon">{icon}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-label">{label}</div>
        </div>"""
    st.markdown(f'<div class="kpi-grid">{cards}</div>', unsafe_allow_html=True)

def severity_badge(severity: str):
    cfg = TIER_CONFIG.get(severity, {"tier": "?", "range": "—", "color": "#8b949e"})
    return (
        f'<span class="severity-badge" '
        f'style="background:rgba({_hex_to_rgb(cfg["color"])},0.12);'
        f'color:{cfg["color"]};border-color:rgba({_hex_to_rgb(cfg["color"])},0.3)">'
        f'Tier {cfg["tier"]} · {severity}</span>'
    )

def _hex_to_rgb(hex_color: str) -> str:
    h = hex_color.lstrip("#")
    return ",".join(str(int(h[i:i+2], 16)) for i in (0, 2, 4))

def _parse_cost(cost_str: str) -> float:
    """Parse '15,234.50 SAR' → 15234.50"""
    return float(cost_str.replace(",", "").replace(" SAR", ""))

# ═══════════════════════════════════════════════════════════════════════════
# PAGE: Dashboard
# ═══════════════════════════════════════════════════════════════════════════
def page_dashboard():
    hero("SmartClaim AI", "AI-powered Arabic vehicle accident cost prediction for the Saudi insurance market")

    metrics = load_metrics()
    if metrics:
        st.markdown('<div class="section-label">Model Performance — Hold-Out Test Set</div>', unsafe_allow_html=True)
        kpi_cards(metrics)
    else:
        st.warning("No metrics found. Run the training pipeline first (`python main.py train`).")

    # Evaluation plots
    plots_dir = "models/plots"
    plot_files = [
        ("actual_vs_predicted.png", "Actual vs Predicted Cost"),
        ("residuals_distribution.png", "Residuals Distribution"),
        ("feature_importance.png", "Top-20 Feature Importance"),
        ("cost_distribution.png", "Cost Distribution"),
    ]
    existing = [(os.path.join(plots_dir, f), t) for f, t in plot_files if os.path.exists(os.path.join(plots_dir, f))]

    if existing:
        st.markdown('<div class="section-label">Diagnostic Plots</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        for i, (path, title) in enumerate(existing):
            with (col1 if i % 2 == 0 else col2):
                with st.container():
                    st.image(path, caption=title, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════
# PAGE: Predict Cost
# ═══════════════════════════════════════════════════════════════════════════
def page_predict():
    hero("🔮 Predict Repair Cost", "Enter an Arabic accident description to get an AI-powered cost estimate")

    # Initialize session state
    if "input_text" not in st.session_state:
        st.session_state.input_text = ""

    # Text area
    text = st.text_area(
        "Arabic Accident Description",
        value=st.session_state.input_text,
        height=140,
        placeholder="...اكتب وصف الحادث هنا، مثال: صدمة في الصدام الخلفي مع خدوش في الرفرف",
        key="accident_text",
    )

    # Example buttons
    st.markdown('<div class="section-label">Quick Examples</div>', unsafe_allow_html=True)
    cols = st.columns(len(EXAMPLE_CASES))
    for i, (label, example_text) in enumerate(EXAMPLE_CASES):
        with cols[i]:
            if st.button(f"💡 {label}", key=f"ex_{i}", use_container_width=True):
                st.session_state.input_text = example_text
                st.rerun()

    st.markdown("")  # spacer

    # Predict button
    predict_clicked = st.button("🔍  Analyze & Predict", type="primary", use_container_width=True)

    # Run prediction
    if predict_clicked and text.strip():
        with st.spinner("🧠 Loading model and generating prediction..."):
            try:
                from src.predict import predict_cost
                from src.feature_extractor import ArabicAccidentFeatureExtractor

                result = predict_cost(text.strip())
                features = ArabicAccidentFeatureExtractor().extract_features(text.strip())
            except Exception as e:
                st.error(f"Prediction failed: {e}")
                return

        if "error" in result:
            st.error(result["error"])
            return

        # ── Result Card ──
        severity = result.get("severity", "Unknown")
        cost_str = result.get("predicted_cost", "0 SAR")
        low_str  = result.get("confidence_range_low", "0 SAR")
        high_str = result.get("confidence_range_high", "0 SAR")

        cost_val = _parse_cost(cost_str)
        low_val  = _parse_cost(low_str)
        high_val = _parse_cost(high_str)

        cfg = TIER_CONFIG.get(severity, {"tier": "?", "range": "—", "color": "#8b949e"})

        st.markdown(
            f"""<div class="result-card">
                <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:16px">
                    <div>
                        {severity_badge(severity)}
                        <div class="cost-display">{cost_val:,.0f} <span class="cost-currency">SAR</span></div>
                        <p style="color:#8b949e;margin:0;font-size:0.9rem">
                            Tier range: {cfg['range']} SAR
                        </p>
                    </div>
                </div>
            </div>""",
            unsafe_allow_html=True,
        )

        # Confidence range bar
        bar_min = low_val * 0.9
        bar_max = high_val * 1.1
        bar_range = bar_max - bar_min if bar_max > bar_min else 1

        fill_left  = ((low_val - bar_min) / bar_range) * 100
        fill_right = ((high_val - bar_min) / bar_range) * 100
        marker_pos = ((cost_val - bar_min) / bar_range) * 100

        st.markdown(
            f"""<div class="confidence-bar-wrapper">
                <div style="font-size:0.82rem;font-weight:600;color:#8b949e;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px">
                    Confidence Range (±15%)
                </div>
                <div class="confidence-bar-track">
                    <div class="confidence-bar-fill" style="left:{fill_left:.1f}%;width:{fill_right-fill_left:.1f}%"></div>
                    <div class="confidence-bar-marker" style="left:{marker_pos:.1f}%"></div>
                </div>
                <div class="confidence-labels">
                    <span>{low_val:,.0f} SAR</span>
                    <span style="color:#00d4aa;font-weight:700">{cost_val:,.0f} SAR</span>
                    <span>{high_val:,.0f} SAR</span>
                </div>
            </div>""",
            unsafe_allow_html=True,
        )

        # Detected features
        st.markdown('<div class="section-label">Detected Features</div>', unsafe_allow_html=True)

        # Determine detected items
        severity_text = "🟢 Very Light" if features["is_very_light"] else \
                        "🟡 Minor" if features["is_minor"] else \
                        "🟠 Moderate" if not features["is_severe"] and not features["is_critical"] else \
                        "🔴 Severe" if features["is_severe"] else "⛔ Critical"
        
        impact_text = "⬆️ Front" if features["is_front"] else \
                      "⬇️ Rear" if features["is_rear"] else \
                      "↔️ Side" if features["is_side"] else "— N/A"

        scratches_text = "✅ Yes" if features["is_scratches"] else "— No"

        items_html = f"""
        <div class="feature-grid">
            <div class="feature-item"><div class="label">Severity</div><div class="value">{severity_text}</div></div>
            <div class="feature-item"><div class="label">Impact Direction</div><div class="value">{impact_text}</div></div>
            <div class="feature-item"><div class="label">Parts Detected</div><div class="value">{features['parts_count']} part{'s' if features['parts_count'] != 1 else ''}</div></div>
            <div class="feature-item"><div class="label">Scratches Only</div><div class="value">{scratches_text}</div></div>
        </div>"""
        st.markdown(items_html, unsafe_allow_html=True)

    elif predict_clicked:
        st.warning("Please enter an Arabic accident description before predicting.")

# ═══════════════════════════════════════════════════════════════════════════
# PAGE: Dataset Explorer
# ═══════════════════════════════════════════════════════════════════════════
def page_dataset():
    hero("📊 Dataset Explorer", "Browse and analyze the SmartClaim training dataset (1,000 Saudi accident records)")

    df = load_dataset()
    if df is None:
        st.error("Dataset not found. Ensure `data/smartclaim_expanded_dataset.xlsx` exists.")
        return

    # Stats row
    stats_html = f"""
    <div class="stat-row">
        <div class="stat-card"><div class="stat-value">{len(df):,}</div><div class="stat-label">Records</div></div>
        <div class="stat-card"><div class="stat-value">{df['cost'].min():,.0f}</div><div class="stat-label">Min Cost (SAR)</div></div>
        <div class="stat-card"><div class="stat-value">{df['cost'].max():,.0f}</div><div class="stat-label">Max Cost (SAR)</div></div>
        <div class="stat-card"><div class="stat-value">{df['cost'].mean():,.0f}</div><div class="stat-label">Mean Cost (SAR)</div></div>
        <div class="stat-card"><div class="stat-value">{df['cost'].std():,.0f}</div><div class="stat-label">Std Dev (SAR)</div></div>
    </div>"""
    st.markdown(stats_html, unsafe_allow_html=True)

    # Interactive charts
    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="section-label">Cost Distribution</div>', unsafe_allow_html=True)
        fig = px.histogram(
            df, x="cost", nbins=40,
            color_discrete_sequence=["#00d4aa"],
            labels={"cost": "Repair Cost (SAR)"},
        )
        fig.update_layout(
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter"),
            margin=dict(l=20, r=20, t=30, b=20),
            xaxis=dict(gridcolor="rgba(255,255,255,0.04)"),
            yaxis=dict(gridcolor="rgba(255,255,255,0.04)", title="Count"),
            bargap=0.06,
        )
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown('<div class="section-label">Text Length Distribution</div>', unsafe_allow_html=True)
        df_temp = df.copy()
        df_temp["text_len"] = df_temp["text"].str.len()
        fig2 = px.histogram(
            df_temp, x="text_len", nbins=30,
            color_discrete_sequence=["#667eea"],
            labels={"text_len": "Characters"},
        )
        fig2.update_layout(
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(family="Inter"),
            margin=dict(l=20, r=20, t=30, b=20),
            xaxis=dict(gridcolor="rgba(255,255,255,0.04)"),
            yaxis=dict(gridcolor="rgba(255,255,255,0.04)", title="Count"),
            bargap=0.06,
        )
        st.plotly_chart(fig2, use_container_width=True)

    # Data table
    st.markdown('<div class="section-label">Dataset Records</div>', unsafe_allow_html=True)
    st.dataframe(
        df[["text", "cost"]].rename(columns={"text": "Arabic Description", "cost": "Repair Cost (SAR)"}),
        use_container_width=True,
        height=420,
    )

# ═══════════════════════════════════════════════════════════════════════════
# PAGE: Model Insights
# ═══════════════════════════════════════════════════════════════════════════
def page_insights():
    hero("🧠 Model Insights", "Understand the SmartClaim AI pipeline architecture and Expert System design")

    # ── Pipeline Architecture ──
    st.markdown('<div class="section-label">End-to-End Pipeline Architecture</div>', unsafe_allow_html=True)

    steps = [
        ("1", "Data Acquisition",       "100 Najm accident reports + 100 Taqdeer cost valuations → augmented to 1,066 records"),
        ("2", "Preprocessing",          "Column standardization, null removal, whitespace normalization → 1,000 clean records"),
        ("3", "Semantic Embedding",     "paraphrase-multilingual-mpnet-base-v2 encodes Arabic text → 768-dim dense vector"),
        ("4", "Arabic Feature Engine",  "Rule-based keyword extraction → 6-dim vector (severity, impact, parts count)"),
        ("5", "Feature Concatenation",  "768-dim embedding + 6-dim features → 774-dim hybrid feature vector"),
        ("6", "Model Training",         "Random Forest vs XGBoost vs Tuned RF — best MAE wins (competitive selection)"),
        ("7", "Expert Cost Engine",     "7-tier rule system constrains ML output to realistic Saudi market cost ranges"),
    ]

    steps_html = ""
    for num, title, desc in steps:
        steps_html += f"""
        <div class="pipeline-step">
            <div class="step-num">{num}</div>
            <div class="step-content">
                <h4>{title}</h4>
                <p>{desc}</p>
            </div>
        </div>"""
    st.markdown(steps_html, unsafe_allow_html=True)

    st.markdown("")

    # ── 7-Tier Expert System ──
    st.markdown('<div class="section-label">7-Tier Expert Cost Estimation Engine</div>', unsafe_allow_html=True)

    tiers_html = ""
    for severity, cfg in TIER_CONFIG.items():
        fill_width = (cfg["tier"] / 7) * 100
        tiers_html += f"""
        <div class="tier-bar">
            <div class="tier-num" style="background:{cfg['color']}">{cfg['tier']}</div>
            <div class="tier-name">{severity}</div>
            <div class="tier-fill" style="background:{cfg['color']};width:{fill_width}%"></div>
            <div class="tier-range">{cfg['range']} SAR</div>
        </div>"""
    st.markdown(tiers_html, unsafe_allow_html=True)

    st.markdown("")

    # ── Feature Importance Plot ──
    fi_path = "models/plots/feature_importance.png"
    if os.path.exists(fi_path):
        st.markdown('<div class="section-label">Feature Importance (Random Forest)</div>', unsafe_allow_html=True)
        st.image(fi_path, use_container_width=True)

    # ── Model Comparison ──
    st.markdown('<div class="section-label">Model Comparison</div>', unsafe_allow_html=True)
    comparison_df = pd.DataFrame({
        "Model": ["Random Forest (Baseline)", "XGBoost", "Tuned Random Forest"],
        "Estimators": [200, 200, "100–500 (search)"],
        "Tuning": ["None", "None", "RandomizedSearchCV (10 iter, 3-fold CV)"],
        "Selection": ["MAE on 80-record validation set"] * 3,
    })
    st.dataframe(comparison_df, use_container_width=True, hide_index=True)

    # ── Embedding Info ──
    st.markdown('<div class="section-label">Embedding Architecture</div>', unsafe_allow_html=True)
    embed_html = """
    <div class="glass-card">
        <h4 style="margin-top:0;color:#00d4aa">paraphrase-multilingual-mpnet-base-v2</h4>
        <p style="color:#8b949e;line-height:1.7;margin-bottom:12px">
            A state-of-the-art multilingual sentence-transformer supporting 50+ languages including Arabic.
            Based on the MPNet backbone with 12 transformer encoder layers and 12 attention heads per layer.
        </p>
        <div class="feature-grid">
            <div class="feature-item"><div class="label">Dense Dims</div><div class="value">768</div></div>
            <div class="feature-item"><div class="label">Rule Features</div><div class="value">6</div></div>
            <div class="feature-item"><div class="label">Total Dims</div><div class="value">774</div></div>
            <div class="feature-item"><div class="label">Languages</div><div class="value">50+</div></div>
        </div>
    </div>"""
    st.markdown(embed_html, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# PAGE: About
# ═══════════════════════════════════════════════════════════════════════════
def page_about():
    hero("ℹ️ About SmartClaim AI", "A university research project for the Saudi vehicle insurance market")

    about_html = """
    <div class="glass-card" style="margin-bottom:24px">
        <h3 style="margin-top:0;color:#e6edf3">SmartClaim AI</h3>
        <p style="color:#c9d1d9;line-height:1.8">
            <strong>SmartClaim AI</strong> is a hybrid Expert-Steered Machine Learning pipeline that processes
            Arabic-language vehicle accident descriptions and delivers high-granularity repair cost estimates
            calibrated to the Saudi automotive repair market.
        </p>
        <p style="color:#8b949e;line-height:1.8">
            The system eliminates the <em>"bucketing bias"</em> problem — where traditional ML models assign
            uniform cost outputs to broad severity categories — by integrating a 7-tier expert rule engine
            that constrains predictions within realistic SAR cost ranges.
        </p>
    </div>"""
    st.markdown(about_html, unsafe_allow_html=True)

    # Project info
    col1, col2 = st.columns(2)
    with col1:
        info_html = """
        <div class="glass-card">
            <h4 style="margin-top:0;color:#00d4aa">📋 Project Details</h4>
            <table style="width:100%;border-collapse:collapse">
                <tr><td style="padding:10px 0;color:#8b949e;width:40%">University</td>
                    <td style="padding:10px 0;color:#e6edf3">King Abdulaziz University</td></tr>
                <tr><td style="padding:10px 0;color:#8b949e;border-top:1px solid rgba(255,255,255,0.06)">Faculty</td>
                    <td style="padding:10px 0;color:#e6edf3;border-top:1px solid rgba(255,255,255,0.06)">Computing & Information Technology</td></tr>
                <tr><td style="padding:10px 0;color:#8b949e;border-top:1px solid rgba(255,255,255,0.06)">Program</td>
                    <td style="padding:10px 0;color:#e6edf3;border-top:1px solid rgba(255,255,255,0.06)">Professional Master in AI</td></tr>
                <tr><td style="padding:10px 0;color:#8b949e;border-top:1px solid rgba(255,255,255,0.06)">Supervisor</td>
                    <td style="padding:10px 0;color:#e6edf3;border-top:1px solid rgba(255,255,255,0.06)">Dr. Somayah Albaradei</td></tr>
                <tr><td style="padding:10px 0;color:#8b949e;border-top:1px solid rgba(255,255,255,0.06)">Author</td>
                    <td style="padding:10px 0;color:#e6edf3;border-top:1px solid rgba(255,255,255,0.06)">Ra'ana Hatim Shaikh</td></tr>
            </table>
        </div>"""
        st.markdown(info_html, unsafe_allow_html=True)

    with col2:
        perf_html = """
        <div class="glass-card">
            <h4 style="margin-top:0;color:#00d4aa">🏆 Key Results</h4>
            <table style="width:100%;border-collapse:collapse">
                <tr><td style="padding:10px 0;color:#8b949e;width:55%">R² Score</td>
                    <td style="padding:10px 0;color:#4ade80;font-weight:700">0.856</td></tr>
                <tr><td style="padding:10px 0;color:#8b949e;border-top:1px solid rgba(255,255,255,0.06)">MAE</td>
                    <td style="padding:10px 0;color:#e6edf3;border-top:1px solid rgba(255,255,255,0.06)">4,661 SAR</td></tr>
                <tr><td style="padding:10px 0;color:#8b949e;border-top:1px solid rgba(255,255,255,0.06)">Dynamic Range</td>
                    <td style="padding:10px 0;color:#e6edf3;border-top:1px solid rgba(255,255,255,0.06)">141× (311 – 43,946 SAR)</td></tr>
                <tr><td style="padding:10px 0;color:#8b949e;border-top:1px solid rgba(255,255,255,0.06)">Dataset Size</td>
                    <td style="padding:10px 0;color:#e6edf3;border-top:1px solid rgba(255,255,255,0.06)">1,000 records</td></tr>
                <tr><td style="padding:10px 0;color:#8b949e;border-top:1px solid rgba(255,255,255,0.06)">Granularity Gain</td>
                    <td style="padding:10px 0;color:#4ade80;font-weight:700;border-top:1px solid rgba(255,255,255,0.06)">47× improvement</td></tr>
            </table>
        </div>"""
        st.markdown(perf_html, unsafe_allow_html=True)

    # Tech stack badges
    st.markdown('<div class="section-label">Technology Stack</div>', unsafe_allow_html=True)
    badges = ["Python 3.10", "Streamlit", "scikit-learn", "XGBoost", "Sentence Transformers",
              "Pandas", "NumPy", "Plotly", "Matplotlib", "Seaborn", "PyTorch", "Joblib"]
    badges_html = '<div class="tech-badge-row">' + "".join(f'<span class="tech-badge">{b}</span>' for b in badges) + '</div>'
    st.markdown(badges_html, unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# Sidebar & Routing
# ═══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(
        """<div style="text-align:center;padding:20px 0 10px 0">
            <div style="font-size:2.4rem">🚗</div>
            <div style="font-size:1.3rem;font-weight:800;background:linear-gradient(135deg,#00d4aa,#667eea);
                -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;margin-top:4px">
                SmartClaim AI
            </div>
            <div style="font-size:0.72rem;color:#484f58;margin-top:4px;letter-spacing:0.06em;text-transform:uppercase">
                Expert-Steered ML Pipeline
            </div>
        </div>""",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    page = st.radio(
        "Navigation",
        ["🏠 Dashboard", "🔮 Predict Cost", "📊 Dataset Explorer", "🧠 Model Insights", "ℹ️ About"],
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown(
        '<div style="text-align:center;font-size:0.7rem;color:#30363d;padding:8px 0">'
        'Built with Streamlit · 2026</div>',
        unsafe_allow_html=True,
    )

# Route
if page == "🏠 Dashboard":
    page_dashboard()
elif page == "🔮 Predict Cost":
    page_predict()
elif page == "📊 Dataset Explorer":
    page_dataset()
elif page == "🧠 Model Insights":
    page_insights()
elif page == "ℹ️ About":
    page_about()
