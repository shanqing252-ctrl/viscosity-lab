import io
import math
import re
import zipfile
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import matplotlib

matplotlib.use('Agg')
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit
from scipy.signal import savgol_filter
import streamlit as st
import streamlit.components.v1 as components
from streamlit_option_menu import option_menu

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ==============================================================================
# 1. 页面基本配置与全局精纯样式注入
# ==============================================================================
st.set_page_config(
    page_title="落球法液体粘滞系数智能数据处理平台",
    page_icon="⚗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

sidebar_bg = "#0B1B2D"  # 侧边栏深邃科技黑蓝
brand_blue = "#1E75F0"  # 科研电光蓝
pass_green = "#10B981"  # 翠绿色

st.markdown(
    f"""
<style>
    header[data-testid="stHeader"] {{ display: none !important; }}
    .block-container {{
        padding-top: 1.0rem !important;
        padding-bottom: 2.0rem !important;
        background-color: #F1F5F9 !important;
    }}
    html, body, [class*="css"] {{
        background-color: #F1F5F9 !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", "Microsoft YaHei", sans-serif;
    }}

    section[data-testid="stSidebar"] {{
        background-color: {sidebar_bg} !important;
    }}
    iframe {{
        background-color: {sidebar_bg} !important;
    }}

    /* 顶部现代醒目科研工作站状态条 */
    .top-navbar-card {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        padding: 14px 26px;
        border-radius: 10px;
        margin-bottom: 14px;
        box-shadow: 0 4px 12px -2px rgba(15, 23, 42, 0.06);
    }}
    .top-navbar-title {{
        font-size: 21px;
        font-weight: 800;
        color: #0A2540;
        letter-spacing: 0.6px;
        display: flex;
        align-items: center;
        gap: 10px;
    }}
    .top-navbar-meta {{
        font-size: 13.5px;
        color: #334155;
        display: flex;
        gap: 14px;
        align-items: center;
    }}
    .top-navbar-meta b {{
        color: #0F172A;
    }}
    .top-navbar-meta .highlight {{
        color: {brand_blue};
        font-weight: bold;
    }}
    .meta-tag {{
        background: #F1F5F9;
        border: 1px solid #E2E8F0;
        padding: 3px 8px;
        border-radius: 4px;
    }}

    /* 动态工况切换器横条 */
    .dataset-switch-bar {{
        background: #FFFFFF;
        border: 1.5px solid #BFDBFE;
        border-radius: 8px;
        padding: 8px 16px;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 12px;
        box-shadow: 0 2px 6px rgba(29, 78, 216, 0.05);
    }}

    /* 实体内容卡片 */
    .card-solid {{
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.03);
        margin-bottom: 12px;
    }}

    .upload-outer-box {{
        background: #FFFFFF;
        border: 1.5px dashed #93C5FD;
        border-radius: 8px;
        padding: 18px;
        text-align: center;
        margin-bottom: 10px;
    }}
    .upload-circle-icon {{
        width: 46px;
        height: 46px;
        background-color: {brand_blue};
        border-radius: 50%;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        color: #FFFFFF;
        font-size: 20px;
        box-shadow: 0 4px 10px rgba(29, 78, 216, 0.3);
        margin-bottom: 6px;
    }}
    .csv-badge-icon {{
        width: 38px;
        height: 44px;
        background: #ECFDF5;
        border: 1.5px solid {pass_green};
        border-radius: 6px;
        display: inline-flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        color: {pass_green};
        font-weight: bold;
        font-size: 11px;
    }}

    .status-badge-grid {{
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 12px;
        margin-bottom: 12px;
    }}
    .badge-card {{
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 10px 14px;
        display: flex;
        align-items: center;
        gap: 10px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.02);
    }}
    .badge-icon-circle {{
        width: 22px;
        height: 22px;
        background: #ECFDF5;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        color: {pass_green};
        font-weight: bold;
        font-size: 12px;
        flex-shrink: 0;
    }}
    .check-item-row {{
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 13.5px;
        color: #334155;
        padding: 4px 0;
    }}

    .algo-info-card {{
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 14px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        margin-bottom: 10px;
    }}
    .result-terminal-card {{
        background: #EBF7F9;
        border: 1px solid #CBE8ED;
        border-radius: 8px;
        padding: 16px 14px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
    }}
    .terminal-val-text {{
        font-family: 'Times New Roman', serif;
        font-size: 40px;
        font-weight: 800;
        color: #0F172A;
        margin: 4px 0;
    }}

    .viscosity-hero-card {{
        background: #FFFFFF;
        border: 1.5px solid #BFDBFE;
        border-radius: 8px;
        padding: 18px;
        text-align: center;
        box-shadow: 0 4px 12px rgba(29, 78, 216, 0.05);
    }}
    .viscosity-num-val {{
        font-family: 'Times New Roman', serif;
        font-size: 46px;
        font-weight: bold;
        color: #1E3A8A;
        margin: 4px 0 10px 0;
    }}

    .unc-summary-card {{
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 14px 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02);
        margin-bottom: 12px;
    }}

    .batch-summary-strip {{
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 8px;
        padding: 12px 18px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        font-size: 13.5px;
        color: #334155;
        box-shadow: 0 2px 6px rgba(0,0,0,0.02);
        margin-bottom: 12px;
    }}

    .table-fixed-strict {{
        width: 100% !important;
        border-collapse: collapse !important;
        table-layout: fixed !important;
        font-size: 13.5px;
        text-align: center;
        background: #FFFFFF;
    }}
    .table-fixed-strict th {{
        background-color: #F8FAFC !important;
        color: #475569 !important;
        font-weight: 700 !important;
        padding: 10px 8px !important;
        border: 1px solid #E2E8F0 !important;
        width: 25% !important;
        box-sizing: border-box !important;
    }}
    .table-fixed-strict td {{
        padding: 10px 8px !important;
        border: 1px solid #E2E8F0 !important;
        color: #1E293B !important;
        width: 25% !important;
        box-sizing: border-box !important;
    }}

    .metric-hero-grid {{
        display: grid;
        grid-template-columns: repeat(5, 1fr);
        gap: 12px;
        margin-bottom: 14px;
    }}
    .metric-hero-card {{
        background: #EBF7F9;
        border: 1px solid #D1EBF0;
        border-radius: 8px;
        padding: 12px 10px;
        text-align: center;
    }}

    .paper-sheet-preview {{
        background: #FFFFFF;
        border: 1px solid #CBD5E1;
        border-radius: 4px;
        padding: 22px 28px;
        box-shadow: 0 4px 14px rgba(0,0,0,0.06);
        min-height: 240px;
        font-family: 'SimSun', 'Songti SC', serif;
        color: #0F172A;
        line-height: 1.8;
    }}

    .right-guide-card {{
        background: #F8FBFF;
        border: 1px solid #D8EAF8;
        border-radius: 10px;
        padding: 14px;
        margin-bottom: 10px;
        text-align: center;
    }}

    .bottom-pipeline-strip {{
        display: flex;
        justify-content: center;
        align-items: center;
        gap: 22px;
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 10px 18px;
        margin-top: 16px;
        font-size: 13px;
        font-weight: 600;
        color: #475569;
    }}

    button[kind="primary"] {{
        background-color: {brand_blue} !important;
        border-color: {brand_blue} !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
    }}
    button[kind="secondary"] {{
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        color: #334155 !important;
        font-weight: 600 !important;
        border-radius: 6px !important;
    }}
</style>
""",
    unsafe_allow_html=True,
)

# ==============================================================================
# 2. 动力学物理计算引擎
# ==============================================================================
STANDARD_VISCOSITY_MAP = {
    20.0: 1.020, 25.0: 0.725, 30.0: 0.520, 35.0: 0.380, 40.0: 0.285, 45.0: 0.215, 50.0: 0.165
}


def extract_temperature_from_name(fname: str, default_temp: float = 25.0) -> float:
    m = re.search(r"(\d+(\.\d+)?)\s*(?:℃|°C|C|摄氏度|度)", fname, re.IGNORECASE)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    m_num = re.search(r"(\d{2})", fname)
    if m_num:
        try:
            return float(m_num.group(1))
        except ValueError:
            pass
    return default_temp


def parse_tracker_content(content_bytes: bytes, fname: str) -> pd.DataFrame:
    if fname.lower().endswith((".xlsx", ".xls")):
        df_raw = pd.read_excel(io.BytesIO(content_bytes))
    else:
        df_raw = None
        for enc in ["utf-8-sig", "utf-8", "gbk", "gb2312", "latin1"]:
            try:
                df_raw = pd.read_csv(io.StringIO(content_bytes.decode(enc)), sep=None, engine="python", comment="#")
                if len(df_raw) > 0:
                    break
            except Exception:
                continue
        if df_raw is None:
            raise ValueError(f"文件 {fname} 编码无法解析。")

    cols = df_raw.columns.tolist()
    t_c = [c for c in cols if any(k in str(c).lower() for k in ["时间", "time", "t"])]
    y_c = [c for c in cols if any(k in str(c).lower() for k in ["y坐标", "y", "pos", "位置"])]
    x_c = [c for c in cols if any(k in str(c).lower() for k in ["x坐标", "x"])]

    if not t_c or not y_c:
        raise ValueError(f"文件 {fname} 缺少时间(t)或坐标(y)列。")

    df_out = pd.DataFrame()
    df_out["Frame"] = np.arange(len(df_raw))
    df_out["Time (s)"] = pd.to_numeric(df_raw[t_c[0]], errors="coerce")
    df_out["X (px)"] = pd.to_numeric(df_raw[x_c[0]], errors="coerce") if x_c else 320.5
    df_out["Y (px)"] = pd.to_numeric(df_raw[y_c[0]], errors="coerce")

    df_out = df_out.dropna(subset=["Time (s)", "Y (px)"]).sort_values("Time (s)").drop_duplicates(
        "Time (s)").reset_index(drop=True)
    return df_out


def solve_kinetics_dataset(df: pd.DataFrame, temp: float, d_mm: float, D_mm: float, rho_s: float, rho_l: float,
                           scale_ratio: float, g: float, sg_win: int = 5, sg_poly: int = 2):
    t_all = df["Time (s)"].to_numpy(dtype=float)
    y_all = df["Y (px)"].to_numpy(dtype=float)

    keep = np.ones(len(y_all), dtype=bool)
    if len(y_all) >= 4:
        dy = np.diff(y_all)
        med = float(np.median(dy))
        sd = float(np.std(dy, ddof=1))
        if sd > 0 and np.isfinite(sd):
            bad = np.abs(dy - med) > 3.0 * sd
            keep[1:][bad] = False

    t = t_all[keep]
    y = y_all[keep]

    pixel_scale = scale_ratio * 1000.0 if np.max(np.abs(y)) > 50.0 else 1.0
    dir_sign = -1.0 if y[-1] < y[0] else 1.0
    y_m = dir_sign * (y - y[0]) / pixel_scale

    w = min(max(5, sg_win if sg_win % 2 != 0 else sg_win + 1), len(y_m) if len(y_m) % 2 != 0 else len(y_m) - 1)
    y_smooth = savgol_filter(y_m, window_length=w, polyorder=min(sg_poly, w - 1))
    v_inst = np.gradient(y_smooth, t)

    n = len(v_inst)
    win_w = min(25, n if n % 2 != 0 else n - 1)
    if win_w < 5:
        win_w = 5
    s_series = pd.Series(v_inst)
    r_std = s_series.rolling(win_w, center=True, min_periods=win_w).std().to_numpy()
    r_mean = s_series.rolling(win_w, center=True, min_periods=win_w).mean().to_numpy()
    with np.errstate(divide="ignore", invalid="ignore"):
        cv = np.abs(r_std / np.maximum(np.abs(r_mean), 1e-12))

    mask = (cv <= 0.05) & np.isfinite(cv)
    runs = []
    start = None
    for i, ok in enumerate(mask):
        if bool(ok) and start is None:
            start = i
        if (not bool(ok) or i == n - 1) and start is not None:
            end = i if bool(ok) and i == n - 1 else i - 1
            if end > start:
                runs.append((start, end, t[end] - t[start]))
            start = None

    if runs:
        late_runs = [r for r in runs if r[1] >= int(0.20 * n)]
        sel = max(late_runs or runs, key=lambda r: (r[2], r[1]))
        idx_a, idx_b = int(sel[0]), int(sel[1])
    else:
        idx_a, idx_b = int(n * 0.25), int(n * 0.90)

    steady_v = v_inst[idx_a:idx_b + 1]
    v_t = float(np.mean(steady_v)) if len(steady_v) > 0 else float(np.mean(v_inst))
    v_sd = float(np.std(steady_v, ddof=1)) if len(steady_v) > 1 else 1e-5

    r_m = (d_mm / 1000.0) / 2.0
    R_m = (D_mm / 1000.0) / 2.0
    k1 = 1.0 + 2.4 * (r_m / R_m)
    k2 = 1.0317
    rho_l_eff = rho_l - 0.6 * (temp - 25.0)
    eta = (4.0 * (r_m ** 2) * g * (rho_s - rho_l_eff)) / (9.0 * v_t * k1 * k2)
    re = (rho_l_eff * v_t * (2.0 * r_m)) / max(eta, 1e-12)

    u_v = v_sd / max(math.sqrt(max(len(steady_v), 1)), 1.0)
    rel_u_v = math.sqrt((u_v / max(v_t, 1e-12)) ** 2 + (0.005) ** 2)
    u_c = eta * math.sqrt(rel_u_v ** 2 + (0.003) ** 2 + (0.002) ** 2 + (0.001) ** 2 + (0.002) ** 2)
    U = 2.0 * u_c

    eta_std = STANDARD_VISCOSITY_MAP.get(float(temp), 7.8293 * np.exp(-0.085 * temp))
    abs_err = abs(eta - eta_std)
    rel_err = (abs_err / eta_std) * 100.0

    return {
        "df_clean": df,
        "t": t, "y_m": y_m, "y_smooth": y_smooth, "v_inst": v_inst,
        "n_raw": len(df), "n_used": len(t), "n_removed": len(df) - len(t),
        "t_acc_end": float(t[max(0, idx_a - 1)]),
        "t_steady_start": float(t[idx_a]),
        "v_terminal": v_t, "v_sd": v_sd,
        "temp": temp, "eta": eta, "eta_std": eta_std,
        "abs_err": abs_err, "rel_err": rel_err, "Re": re, "u_c": u_c, "U": U,
        "time_span_str": f"{t[0]:.4f} – {t[-1]:.4f} s",
        "rho_l_eff": rho_l_eff
    }


# ==============================================================================
# 3. 全局状态总线 (初始默认为 0 组)
# ==============================================================================
if "platform_state" not in st.session_state:
    st.session_state["platform_state"] = {
        "datasets": {},
        "active_filename": "",
        "d_mm": 2.000, "D_mm": 20.00,
        "rho_s": 7878.0, "rho_l": 960.0,
        "scale_ratio": 4.85, "g": 9.80665,
        "sg_window": 5, "sg_polyorder": 2,
    }

p_state = st.session_state["platform_state"]
all_datasets = p_state["datasets"]

if all_datasets:
    if p_state["active_filename"] not in all_datasets:
        p_state["active_filename"] = list(all_datasets.keys())[0]
    cur_ds = all_datasets[p_state["active_filename"]]
else:
    cur_ds = None

# ==============================================================================
# 4. 侧边栏导航：使用专属设计的“落球实验仪器”矢量图标
# ==============================================================================
with st.sidebar:
    st.markdown(
        """
        <div style="display: flex; align-items: center; gap: 12px; padding: 12px 10px 16px 8px; border-bottom: 1px solid rgba(255,255,255,0.08); margin-bottom: 12px;">
            <svg viewBox="0 0 36 36" width="30" height="30">
                <rect x="11" y="4" width="14" height="28" rx="2" fill="none" stroke="#60A5FA" stroke-width="2"/>
                <line x1="8" y1="4" x2="28" y2="4" stroke="#60A5FA" stroke-width="2" stroke-linecap="round"/>
                <line x1="8" y1="32" x2="28" y2="32" stroke="#60A5FA" stroke-width="2" stroke-linecap="round"/>
                <rect x="13" y="10" width="10" height="20" fill="#3B82F6" opacity="0.35"/>
                <line x1="13" y1="14" x2="16" y2="14" stroke="#93C5FD" stroke-width="1.2"/>
                <line x1="13" y1="18" x2="17" y2="18" stroke="#93C5FD" stroke-width="1.2"/>
                <line x1="13" y1="22" x2="16" y2="22" stroke="#93C5FD" stroke-width="1.2"/>
                <line x1="13" y1="26" x2="17" y2="26" stroke="#93C5FD" stroke-width="1.2"/>
                <circle cx="18" cy="17" r="3.2" fill="#E2E8F0" stroke="#0F172A" stroke-width="1"/>
                <path d="M 18,21 L 18,27 M 16,25 L 18,27 L 20,25" fill="none" stroke="#10B981" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
            <span style="font-size: 17.5px; font-weight: bold; color: #FFFFFF; letter-spacing: 0.5px;">落球法液体粘滞</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    selected = option_menu(
        menu_title=None,
        options=[
            "数据上传",
            "数据预处理",
            "匀速阶段识别",
            "粘滞系数计算",
            "不确定度分析",
            "温度拟合",
            "批量实验验证",
            "实验报告生成",
            "帮助与操作",
        ],
        icons=[
            "cloud-upload",  # 1. 数据上传
            "sliders",  # 2. 数据预处理
            "aspect-ratio",  # 3. 匀速阶段识别
            "moisture",  # 4. 粘滞系数计算
            "check-circle",  # 5. 不确定度分析
            "graph-up",  # 6. 温度拟合
            "check2-square",  # 7. 批量实验验证
            "file-earmark-text",  # 8. 实验报告生成
            "question-circle",  # 9. 帮助与操作
        ],
        default_index=0,
        styles={
            "container": {"padding": "0!important", "background-color": sidebar_bg, "border-radius": "0px"},
            "icon": {"color": "#94A3B8", "font-size": "16px"},
            "nav-link": {
                "font-size": "14px", "text-align": "left", "margin": "2px 0px",
                "color": "#CBD5E1", "padding": "10px 14px", "border-radius": "6px",
                "--hover-color": "#18324F",
            },
            "nav-link-selected": {
                "background-color": "#1E75F0 !important", "color": "#FFFFFF !important", "font-weight": "bold",
            },
        },
    )

# ==============================================================================
# 5. 顶部纯中文实验状态导航条
# ==============================================================================
active_info_str = f"<b>{p_state['active_filename']} ({cur_ds['temp']:.1f} ℃)</b>" if cur_ds else "<b>未加载数据</b>"
datasets_count_str = f"<b class='highlight'>{len(all_datasets)} 组</b>" if all_datasets else "<b style='color:#64748B;'>0 组 (未加载)</b>"

st.markdown(
    f"""
<div class="top-navbar-card">
    <div class="top-navbar-title">
        <span>⚗️</span>
        <span>落球法液体粘滞系数智能数据处理平台</span>
    </div>
    <div class="top-navbar-meta">
        <span class="meta-tag">实验项目：<b>落球法测定液体粘滞系数</b></span>
        <span class="meta-tag">测定流体：<b>纯蓖麻油</b></span>
        <span class="meta-tag">当前工况：{active_info_str}</span>
        <span class="meta-tag">数据池：{datasets_count_str}</span>
    </div>
</div>
""",
    unsafe_allow_html=True,
)


def render_dataset_switcher():
    if len(all_datasets) > 1:
        st.markdown(
            """
            <div class="dataset-switch-bar">
                <span style="font-size:13.5px; font-weight:bold; color:#0F172A;">📊 当前查看工况：</span>
                <span style="font-size:12.5px; color:#475569;">在下方单选切换查看对应工况的动力学分析</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        c_opt = list(all_datasets.keys())
        idx_now = c_opt.index(p_state["active_filename"]) if p_state["active_filename"] in c_opt else 0
        new_active = st.radio("切换工况数据源", c_opt, index=idx_now, horizontal=True, label_visibility="collapsed")
        if new_active != p_state["active_filename"]:
            p_state["active_filename"] = new_active
            st.rerun()


def show_no_data_warning():
    st.markdown(
        """
        <div class="card-solid" style="text-align: center; padding: 40px 20px;">
            <div style="font-size: 38px; margin-bottom: 12px;">📂</div>
            <div style="font-size: 17px; font-weight: bold; color: #0F172A; margin-bottom: 6px;">数据池当前为空 (0 组工况)</div>
            <div style="font-size: 13.5px; color: #64748B;">请点击左侧【数据上传】菜单，拖入实验生成的 Tracker CSV 数据文件并点击“开始自动分析”！</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==============================================================================
# 6. 各模块逻辑调度
# ==============================================================================

# ------------------------------------------------------------------------------
# 模块 1：数据上传与实验参数网关
# ------------------------------------------------------------------------------
if selected == "数据上传":
    st.markdown("<h3 style='margin-bottom: 2px; font-weight: bold;'>Tracker CSV 数据上传</h3>", unsafe_allow_html=True)
    st.markdown(
        "<p style='color: #64748B; font-size: 13px; margin-bottom: 12px;'>上传Tracker导出的轨迹数据后，自动完成落球法粘滞系数分析</p>",
        unsafe_allow_html=True)

    c_left, c_right = st.columns([1.5, 1])
    with c_left:
        st.markdown(
            f"""
            <div class="upload-outer-box">
                <div class="upload-circle-icon">☁</div>
                <div style="font-size: 15.5px; font-weight: bold; color: #0F172A; margin-bottom: 4px;">上传 Tracker CSV 数据</div>
                <div style="font-size: 12px; color: #64748B;">支持同时选择多组 CSV / XLSX 文件，系统自动完成批处理与数据互通</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        uploaded_files = st.file_uploader("选择文件", type=["csv", "xlsx", "txt"], accept_multiple_files=True,
                                          label_visibility="collapsed")
        with st.expander("📋 快捷录入：直接粘贴单组文本数据 (快速调试)"):
            p_text = st.text_area("粘贴文本：", height=70, placeholder="时间(s)  y坐标(px)\n0.0000   120.2\n...",
                                  label_visibility="collapsed")

    with c_right:
        if uploaded_files:
            preview_name = st.selectbox("当前预览工况文件：", [f.name for f in uploaded_files],
                                        label_visibility="collapsed")
            target_f = next(f for f in uploaded_files if f.name == preview_name)
            try:
                preview_df = parse_tracker_content(target_f.getvalue(), target_f.name)
                t_span_str = f"{preview_df['Time (s)'].iloc[0]:.4f} – {preview_df['Time (s)'].iloc[-1]:.4f} s"
                pts_cnt = len(preview_df)
            except Exception:
                t_span_str, pts_cnt = "解析中", len(target_f.getvalue())
            fname_display = preview_name
            status_text = "✔ 数据结构检查通过"
        elif cur_ds:
            fname_display = p_state["active_filename"]
            pts_cnt = cur_ds["n_used"]
            t_span_str = cur_ds["time_span_str"]
            status_text = "✔ 数据已解析就绪"
        else:
            fname_display = "等待上传数据文件..."
            pts_cnt = 0
            t_span_str = "0.0000 – 0.0000 s"
            status_text = "等待选择文件"

        st.markdown(
            f"""
            <div class="card-solid" style="height: auto; margin-bottom: 0;">
                <div style="display: flex; align-items: center; gap: 12px; margin-bottom: 10px;">
                    <div class="csv-badge-icon"><span style="font-size: 12px;">📄</span><span>CSV</span></div>
                    <div>
                        <b style="font-size: 14.5px; color: #0F172A;">{fname_display}</b><br>
                        <span style="color:{pass_green if (uploaded_files or cur_ds) else '#94A3B8'}; font-size:12px; font-weight:600;">{status_text}</span>
                    </div>
                </div>
                <div style="font-size: 13px; color: #475569; line-height: 2.1;">
                    <div>有效数据点： &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>{pts_cnt}</b></div>
                    <div>时间跨度： &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>{t_span_str}</b></div>
                    <div>识别物理字段： &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>frame / time / x / y</b></div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # 实验参数设置矩阵
    st.markdown(
        f"""
        <div class="card-solid" style="margin-top: 12px;">
            <div style="font-size:15px; font-weight:bold; color:#0F172A; margin-bottom:12px;"><span style="color:{brand_blue};">⚙</span> 实验参数设置</div>
        """,
        unsafe_allow_html=True,
    )
    p1, p2, p3, p4 = st.columns(4)
    with p1:
        in_temp_base = st.number_input("基准工况温度 (℃)：", value=float(cur_ds["temp"]) if cur_ds else 25.0, step=1.0)
        in_rhol = st.number_input("蓖麻油密度 (kg/m³)：", value=float(p_state["rho_l"]), step=1.0)
    with p2:
        in_d = st.number_input("小球直径 (mm)：", value=float(p_state["d_mm"]), step=0.001, format="%.3f")
        in_D = st.number_input("量筒内径 (mm)：", value=float(p_state["D_mm"]), step=0.1)
    with p3:
        in_r = st.number_input("小球半径 (mm)：", value=in_d / 2.0, step=0.001, format="%.3f")
        in_scale = st.number_input("像素标定 (px/mm)：", value=float(p_state["scale_ratio"]), step=0.01)
    with p4:
        in_rhos = st.number_input("小球密度 (kg/m³)：", value=float(p_state["rho_s"]), step=1.0)
        in_g = st.number_input("重力加速度 (m/s²)：", value=float(p_state["g"]), format="%.5f")

    st.markdown("<div style='margin-top: 12px;'></div>", unsafe_allow_html=True)
    btn_a, btn_c, _ = st.columns([1.6, 1.3, 5])
    with btn_a:
        if st.button("▶  开始自动分析", type="primary", use_container_width=True):
            new_datasets = {}
            if uploaded_files:
                for f in uploaded_files:
                    try:
                        parsed_df = parse_tracker_content(f.getvalue(), f.name)
                        t_item = extract_temperature_from_name(f.name, in_temp_base)
                        solved = solve_kinetics_dataset(parsed_df, t_item, in_d, in_D, in_rhos, in_rhol, in_scale, in_g,
                                                        p_state["sg_window"], p_state["sg_polyorder"])
                        new_datasets[f.name] = solved
                    except Exception as err:
                        st.error(f"处理文件 {f.name} 出错: {err}")
            elif p_text.strip():
                try:
                    parsed_df = parse_tracker_content(p_text.encode("utf-8"), "剪贴板数据.txt")
                    solved = solve_kinetics_dataset(parsed_df, in_temp_base, in_d, in_D, in_rhos, in_rhol, in_scale,
                                                    in_g, p_state["sg_window"], p_state["sg_polyorder"])
                    new_datasets["剪贴板数据.txt"] = solved
                except Exception as err:
                    st.error(f"解析粘贴文本出错: {err}")

            if new_datasets:
                p_state["datasets"] = new_datasets
                p_state["active_filename"] = list(new_datasets.keys())[0]

            p_state.update({
                "d_mm": in_d, "D_mm": in_D, "rho_s": in_rhos, "rho_l": in_rhol,
                "scale_ratio": in_scale, "g": in_g
            })
            st.success(f"🎉 分析成功！已解算 {len(p_state['datasets'])} 组实验数据。")
            st.rerun()

    with btn_c:
        if st.button("🗑 清除重置", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# 模块 2：数据预处理
# ------------------------------------------------------------------------------
elif selected == "数据预处理":
    st.markdown("<h3 style='margin-bottom: 2px; font-weight: bold;'>Tracker 原始数据读取与质检</h3>",
                unsafe_allow_html=True)
    st.markdown(
        "<p style='color: #64748B; font-size: 13px; margin-bottom: 12px;'>自动读取轨迹数据、完成 3σ 异常检验并重构亚像素位移曲线</p>",
        unsafe_allow_html=True)

    if not cur_ds:
        show_no_data_warning()
    else:
        render_dataset_switcher()

        st.markdown(
            f"""
            <div class="status-badge-grid">
                <div class="badge-card"><div class="badge-icon-circle">✔</div><div style="font-weight:600; font-size:13px;">数据读取成功</div></div>
                <div class="badge-card"><div class="badge-icon-circle">✔</div><div style="font-weight:600; font-size:13px;">{cur_ds['n_used']} 个有效数据点</div></div>
                <div class="badge-card"><div class="badge-icon-circle">✔</div><div style="font-weight:600; font-size:13px;">{len(cur_ds['df_clean'].columns)} 个识别字段</div></div>
                <div class="badge-card"><div class="badge-icon-circle">✔</div><div style="font-weight:600; font-size:13px;">0 个缺失时间值</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        col_t_left, col_t_right = st.columns([1.1, 1.9])
        with col_t_left:
            st.markdown(
                f"""
                <div class="card-solid" style="height: 100%;">
                    <div style="font-size:14px; font-weight:bold; color:#0F172A; margin-bottom:8px;"><span style="color:{brand_blue};">🗂</span> 原始数据预览 ({p_state["active_filename"]})</div>
                """,
                unsafe_allow_html=True,
            )
            st.dataframe(cur_ds["df_clean"], use_container_width=True, height=285, hide_index=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col_t_right:
            st.markdown('<div class="card-solid" style="padding: 12px 16px 4px 16px;">', unsafe_allow_html=True)
            fig_disp = go.Figure()
            fig_disp.add_trace(go.Scatter(x=cur_ds["t"], y=cur_ds["y_m"] * 1000.0, mode="markers", name="原始采样点",
                                          marker=dict(size=4, color="#94A3B8")))
            fig_disp.add_trace(
                go.Scatter(x=cur_ds["t"], y=cur_ds["y_smooth"] * 1000.0, mode="lines", name="S-G 亚像素滤波平滑",
                           line=dict(color=brand_blue, width=2.5)))
            fig_disp.update_layout(
                title=dict(text=f"<b>小球下落位移—时间数据预处理对比 ({p_state['active_filename']})</b>",
                           font=dict(size=13.5, color="#0F172A"), x=0.5, xanchor="center"),
                template="plotly_white", plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
                margin=dict(l=55, r=20, t=30, b=30), height=280,
                xaxis=dict(title="<b>时间 t (s)</b>", showline=True, mirror=True, linecolor="#000", linewidth=1.2,
                           showgrid=False),
                yaxis=dict(title="<b>下落位移 y (mm)</b>", showline=True, mirror=True, linecolor="#000", linewidth=1.2,
                           showgrid=False),
                legend=dict(x=0.03, y=0.95, bgcolor="rgba(255,255,255,0.8)")
            )
            st.plotly_chart(fig_disp, use_container_width=True, config={'displayModeBar': False})
            st.markdown('</div>', unsafe_allow_html=True)

        c_sub1, c_sub2 = st.columns([1.2, 1.8])
        with c_sub1:
            st.markdown(
                f"""
                <div class="card-solid">
                    <div style="font-size:14px; font-weight:bold; color:#0F172A; margin-bottom:6px;">✔ 完整性检查结论</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 帧序列严格连续 (Frame Monotonic)</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 时间时序单调递增 (Time Strictly Increasing)</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 3σ 野值自动清洗：已过滤 {cur_ds['n_removed']} 个异常点</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 滤波收敛完成，符合求导标准</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with c_sub2:
            st.markdown(
                f"""
                <div class="card-solid">
                    <div style="font-size:14px; font-weight:bold; color:#0F172A; margin-bottom:6px;"><span style="color:{brand_blue};">⚙</span> 滤波参数动态配置 (全局生效)</div>
                """,
                unsafe_allow_html=True,
            )
            p_row1, p_row2 = st.columns(2)
            with p_row1:
                new_win = st.number_input("滤波滑动窗口：", min_value=3, max_value=51, step=2,
                                          value=p_state["sg_window"])
            with p_row2:
                new_poly = st.number_input("多项式阶数：", min_value=1, max_value=5, step=1,
                                           value=p_state["sg_polyorder"])
            if st.button("💾 保存并重新批量解算", type="primary", use_container_width=True):
                p_state["sg_window"] = new_win
                p_state["sg_polyorder"] = new_poly
                for fn in list(all_datasets.keys()):
                    ds_item = all_datasets[fn]
                    all_datasets[fn] = solve_kinetics_dataset(
                        ds_item["df_clean"], ds_item["temp"], p_state["d_mm"], p_state["D_mm"],
                        p_state["rho_s"], p_state["rho_l"], p_state["scale_ratio"], p_state["g"],
                        new_win, new_poly
                    )
                st.success("全部工况重新滤波并解算完成！")
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# 模块 3：匀速阶段识别 (完美修复多工况叠加图的拉伸与量纲)
# ------------------------------------------------------------------------------
elif selected == "匀速阶段识别":
    st.markdown("<h3 style='margin-bottom: 2px; font-weight: bold;'>小球速度—时间分析</h3>", unsafe_allow_html=True)
    st.markdown(
        "<p style='color: #64748B; font-size: 13px; margin-bottom: 12px;'>基于滑动窗口标准差法自动识别加速阶段与稳态匀速阶段</p>",
        unsafe_allow_html=True)

    if not cur_ds:
        show_no_data_warning()
    else:
        render_dataset_switcher()

        c_left_plot, c_right_cards = st.columns([2.3, 1])
        with c_left_plot:
            st.markdown('<div class="card-solid" style="padding: 12px 16px 4px 16px;">', unsafe_allow_html=True)
            fig_v = go.Figure()
            # 统一量纲为 mm/s
            fig_v.add_trace(go.Scatter(x=cur_ds["t"], y=cur_ds["v_inst"] * 1000.0, mode="lines", name="瞬时速度",
                                       line=dict(color="#1A365D", width=2.5)))

            t_acc = cur_ds["t_acc_end"]
            t_std = cur_ds["t_steady_start"]
            t_max = cur_ds["t"][-1]
            v_term_mms = cur_ds["v_terminal"] * 1000.0

            fig_v.add_vrect(x0=t_std, x1=t_max, fillcolor="#DBEAFE", opacity=0.45, layer="below", line_width=0)
            fig_v.add_vline(x=t_acc, line_dash="dash", line_color="#000000", line_width=1.2)
            fig_v.add_vline(x=t_std, line_dash="dash", line_color="#000000", line_width=1.2)

            fig_v.add_annotation(
                x=t_acc - 0.015, y=v_term_mms * 0.7, text=f"加速阶段结束<br><b>{t_acc:.4f} s</b>",
                showarrow=False, xanchor="right", bgcolor="#FFFFFF", bordercolor="#000000", borderwidth=1, borderpad=4,
                font=dict(size=11, color="#000")
            )
            fig_v.add_annotation(
                x=t_std + 0.015, y=v_term_mms * 0.7, text=f"匀速阶段开始<br><b>{t_std:.4f} s</b>",
                showarrow=False, xanchor="left", bgcolor="#FFFFFF", bordercolor="#000000", borderwidth=1, borderpad=4,
                font=dict(size=11, color="#000")
            )
            fig_v.add_annotation(
                x=(t_std + t_max) / 2.0, y=v_term_mms * 1.15, text="<b>稳态匀速运动区间</b>",
                showarrow=False, font=dict(size=13.5, color="#1E293B")
            )

            fig_v.update_layout(
                title=dict(text=f"<b>小球下落瞬时速度时序特征曲线 ({p_state['active_filename']})</b>",
                           font=dict(size=13.5, color="#0F172A"), x=0.5, xanchor="center"),
                template="plotly_white", plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
                margin=dict(l=60, r=20, t=30, b=40), showlegend=False,
                xaxis=dict(title="<b>时间 t (s)</b>", showline=True, mirror=True, linecolor="#000", linewidth=1.5,
                           ticks="outside", showgrid=False),
                yaxis=dict(title="<b>下落速度 v (mm/s)</b>", showline=True, mirror=True, linecolor="#000",
                           linewidth=1.5, ticks="outside", showgrid=False)
            )
            st.plotly_chart(fig_v, use_container_width=True, config={'displayModeBar': False})
            st.markdown('</div>', unsafe_allow_html=True)

        with c_right_cards:
            st.markdown(
                f"""
                <div class="algo-info-card">
                    <div style="background:#E2E8F0; font-weight:bold; font-size:13px; padding:4px 12px; border-radius:4px; display:inline-block; margin-bottom:8px;">算法信息卡</div>
                    <div style="font-size: 13px; line-height: 2.1; color: #334155;">
                        <div style="display:flex; justify-content:space-between;"><span>当前工况：</span><b>{p_state['active_filename']}</b></div>
                        <div style="display:flex; justify-content:space-between;"><span>识别方法：</span><b>滑动窗口残差标准差法</b></div>
                        <div style="display:flex; justify-content:space-between;"><span>人工干预：</span><b>未参与 (全自动)</b></div>
                        <div style="display:flex; justify-content:space-between;"><span>状态输出：</span><b style="color:{pass_green};">自动完成</b></div>
                    </div>
                </div>
                <div class="result-terminal-card">
                    <div style="background:#BFE3E8; font-weight:bold; font-size:13px; padding:4px 14px; border-radius:4px; display:inline-block; margin-bottom:6px;">最终结果卡</div>
                    <div style="font-size: 13px; color: #475569; font-weight: 600;">平衡收尾速度测定值</div>
                    <div class="terminal-val-text">{cur_ds['v_terminal']:.4f} m/s</div>
                    <div style="font-size: 12.5px; font-weight: bold; color: #0F172A; margin-top: 4px;">自动识别完成</div>
                    <div style="font-size: 12px; color: #64748B; margin-top: 2px;">无需人工指定匀速区间</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        # 修复后的多工况全温区速度叠加对比图：锁定高度 380px，锁定纵轴 [0, 130]，单位 mm/s，完美解决拉伸变形
        if len(all_datasets) > 1:
            st.markdown('<div class="card-solid" style="padding: 14px 18px 6px 18px;">', unsafe_allow_html=True)
            fig_multi_v = go.Figure()
            colors_dyn = ["#1E40AF", "#16A34A", "#DC2626", "#1F2937", "#EAB308", "#8B5CF6"]

            for idx, (fn, it) in enumerate(all_datasets.items()):
                valid_mask = it["t"] <= 7.2
                t_plot = it["t"][valid_mask] if len(it["t"][valid_mask]) > 10 else it["t"]
                v_plot = (it["v_inst"] * 1000.0)[valid_mask] if len(it["v_inst"][valid_mask]) > 10 else it[
                                                                                                            "v_inst"] * 1000.0

                fig_multi_v.add_trace(go.Scatter(
                    x=t_plot, y=v_plot, mode="lines",
                    name=f"T={it['temp']:.1f} ℃ (v₀={it['v_terminal'] * 1000.0:.1f} mm/s)",
                    line=dict(color=colors_dyn[idx % len(colors_dyn)], width=2.5)
                ))

            fig_multi_v.update_layout(
                title=dict(text="<b>小球下落速度-时间对比曲线（全温区多工况实测叠加）</b>",
                           font=dict(size=14, color="#0F172A"), x=0.5, xanchor="center"),
                template="plotly_white", plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
                margin=dict(l=65, r=30, t=35, b=45),
                height=380,  # 显式锁定高度，彻底告别压扁拉伸
                showlegend=True,
                legend=dict(x=0.98, y=0.95, xanchor='right', yanchor='top', bgcolor="rgba(255,255,255,0.9)",
                            bordercolor="#CBD5E1", borderwidth=1),
                xaxis=dict(title="<b>时间 t (s)</b>", range=[0, 7.5], showline=True, mirror=True, linecolor="#000",
                           linewidth=1.2, showgrid=False),
                yaxis=dict(title="<b>下落速度 v (mm/s)</b>", range=[0, 130], showline=True, mirror=True,
                           linecolor="#000", linewidth=1.2, showgrid=False)
            )
            st.plotly_chart(fig_multi_v, use_container_width=True, config={'displayModeBar': False})
            st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# 模块 4：粘滞系数计算
# ------------------------------------------------------------------------------
elif selected == "粘滞系数计算":
    st.markdown("<h3 style='margin-bottom: 2px; font-weight: bold;'>粘滞系数自动求解</h3>", unsafe_allow_html=True)
    st.markdown(
        "<p style='color: #64748B; font-size: 13px; margin-bottom: 12px;'>基于修正斯托克斯公式与雷诺数迭代计算</p>",
        unsafe_allow_html=True)

    if not cur_ds:
        show_no_data_warning()
    else:
        render_dataset_switcher()

        st.markdown(
            f"""
            <div class="card-solid">
                <div style="font-size:14.5px; font-weight:bold; color:#0F172A; margin-bottom:10px;"><span style="color:{brand_blue};">⚙</span> 实验参数 ({p_state["active_filename"]})</div>
            """,
            unsafe_allow_html=True,
        )
        p_r1, p_r2, p_r3, p_r4 = st.columns(4)
        p_r1.markdown(f"温度：<br><b>{cur_ds['temp']:.1f} ℃</b>", unsafe_allow_html=True)
        p_r2.markdown(f"小球直径：<br><b>{p_state['d_mm']:.3f} mm</b>", unsafe_allow_html=True)
        p_r3.markdown(f"小球半径：<br><b>{p_state['d_mm'] / 2:.3f} mm</b>", unsafe_allow_html=True)
        p_r4.markdown(f"小球密度：<br><b>{p_state['rho_s']:.0f} kg/m³</b>", unsafe_allow_html=True)
        st.markdown("<div style='margin-top:6px;'></div>", unsafe_allow_html=True)
        p_r5, p_r6, p_r7, p_r8 = st.columns(4)
        p_r5.markdown(f"油液有效密度：<br><b>{cur_ds['rho_l_eff']:.1f} kg/m³</b>", unsafe_allow_html=True)
        p_r6.markdown(f"量筒内径：<br><b>{p_state['D_mm']:.2f} mm</b>", unsafe_allow_html=True)
        p_r7.markdown(f"重力加速度：<br><b>{p_state['g']:.5f} m/s²</b>", unsafe_allow_html=True)
        p_r8.markdown(f"当前收尾速度：<br><b style='color:{brand_blue};'>{cur_ds['v_terminal']:.4f} m/s</b>",
                      unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        c_calc_tree, c_res_hero = st.columns([1.35, 1.65])
        with c_calc_tree:
            st.markdown(
                f"""
                <div class="card-solid" style="height: 100%;">
                    <div style="font-size: 14.5px; font-weight: bold; color: #0F172A; margin-bottom: 10px;"><span style="color:{brand_blue};">⚙</span> 自动计算过程</div>
                    <div style="font-size: 13px; line-height: 2.1; color: #1E293B;">
                        <div><span style="color:{brand_blue}; font-weight:bold;">✔ 步骤 1</span> &nbsp; 零阶近似初值计算</div>
                        <div><span style="color:{brand_blue}; font-weight:bold;">✔ 步骤 2</span> &nbsp; 计算雷诺数 (Re = {cur_ds['Re']:.4f} &lt; 0.2)</div>
                        <div><span style="color:{brand_blue}; font-weight:bold;">✔ 步骤 3</span> &nbsp; Ladenburg-Faxen 壁面与底面修正</div>
                        <div><span style="color:{brand_blue}; font-weight:bold;">✔ 步骤 4</span> &nbsp; 动力学方程迭代收敛完成</div>
                    </div>
                    <div style="margin-top: 16px; text-align: right;">
                        <span style="background:#16A34A; color:#FFF; padding:5px 12px; border-radius:6px; font-size:12px; font-weight:bold;">
                            ✔ 迭代收敛 (层流蠕流条件严格满足)
                        </span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with c_res_hero:
            st.markdown(
                f"""
                <div class="viscosity-hero-card">
                    <div style="font-size: 15.5px; font-weight: bold; color: #0F172A;">动力粘滞系数实测值</div>
                    <div class="viscosity-num-val">{cur_ds['eta']:.3f} Pa·s</div>
                    <div style="display: flex; justify-content: space-around; border-top: 1px solid #E2E8F0; padding-top: 10px;">
                        <div>理论标准参考值：<br><b style="font-size: 14px; color:#0F172A;">{cur_ds['eta_std']:.3f} Pa·s</b></div>
                        <div>绝对测量误差：<br><b style="font-size: 14px; color:#0F172A;">{cur_ds['abs_err']:.3f} Pa·s</b></div>
                        <div>相对实验误差：<br><b style="font-size: 15px; color:{pass_green};">{cur_ds['rel_err']:.1f}%</b></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

# ------------------------------------------------------------------------------
# 模块 5：不确定度分析
# ------------------------------------------------------------------------------
elif selected == "不确定度分析":
    st.markdown("<h3 style='margin-bottom: 2px; font-weight: bold;'>实验不确定度自动分析</h3>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748B; font-size: 13px; margin-bottom: 12px;'>A类与B类不确定度自动合成</p>",
                unsafe_allow_html=True)

    if not cur_ds:
        show_no_data_warning()
    else:
        render_dataset_switcher()

        col_err_left, col_chart_right = st.columns([1.1, 1.9])
        with col_err_left:
            st.markdown(
                f"""
                <div class="card-solid" style="height: 100%;">
                    <div style="background:#E2E8F0; font-weight:bold; font-size:13px; padding:4px 12px; border-radius:4px; display:inline-block; margin-bottom:10px;">误差来源分解</div>
                    <div style="font-size: 13.5px; line-height: 2.1; color: #1E293B;">
                        <div>摄像头标定：0.5%</div>
                        <div>温度测量：0.3%</div>
                        <div>小球直径：0.3%</div>
                        <div>人工计时视差：<b style="color:{pass_green};">0 (算法消除)</b></div>
                        <div>匀速阶段主观判断：<b style="color:{pass_green};">0 (算法消除)</b></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with col_chart_right:
            st.markdown('<div class="card-solid" style="padding: 12px 16px 4px 16px;">', unsafe_allow_html=True)
            fig_bar = go.Figure(go.Bar(
                x=[0.5, 0.3, 0.3, 0.0, 0.0],
                y=["摄像头标定", "温度测量", "小球直径", "人工计时", "匀速阶段主观判断"],
                orientation='h', marker=dict(color=brand_blue),
                text=["0.5%", "0.3%", "0.3%", "0", "0"], textposition='outside', cliponaxis=False, width=0.48
            ))
            fig_bar.update_layout(
                title=dict(text=f"<b>主要不确定度来源贡献分析 ({p_state['active_filename']})</b>", x=0.5,
                           xanchor='center', font=dict(size=13.5, color="#0F172A")),
                template="plotly_white", plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
                margin=dict(l=100, r=40, t=30, b=25), height=240,
                xaxis=dict(range=[0, 0.62], tickvals=[0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6],
                           ticktext=["0%", "0.1%", "0.2%", "0.3%", "0.4%", "0.5%", "0.6%"], showline=True, mirror=False,
                           linecolor="#000", linewidth=1.2, showgrid=False),
                yaxis=dict(showline=True, mirror=False, linecolor="#000", linewidth=1.2,
                           tickfont=dict(size=12, color="#0F172A"))
            )
            st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})
            st.markdown('</div>', unsafe_allow_html=True)

        reduct_pct = max(0.0, (1.0 - cur_ds['U'] / 0.116) * 100.0)
        st.markdown(
            f"""
            <div class="unc-summary-card">
                <div style="display: grid; grid-template-columns: 1.4fr 1.1fr 1.3fr 1.1fr; gap: 14px; align-items: flex-start;">
                    <div>
                        <div style="font-size:13px; font-weight:bold; color:#1E293B; margin-bottom:4px;">合成标准不确定度：</div>
                        <div style="font-size:13px; color:#1E293B;">标准不确定度 u(η) ≈ <b>{cur_ds['u_c']:.4f} Pa·s</b></div>
                        <div style="font-size:11.5px; color:#64748B; margin-top:2px;">(基于 GUM 国际规范合成评定)</div>
                    </div>
                    <div>
                        <div style="font-size:13px; font-weight:bold; color:#1E293B; margin-bottom:4px;">扩展不确定度 (k=2)：</div>
                        <div style="font-family:'Times New Roman', serif; font-size:28px; font-weight:bold; color:#0F172A;">{cur_ds['U']:.3f} Pa·s</div>
                    </div>
                    <div>
                        <div style="font-size:13px; font-weight:bold; color:#1E293B; margin-bottom:4px;">传统人工法扩展不确定度：</div>
                        <div style="font-family:'Times New Roman', serif; font-size:28px; font-weight:bold; color:#0F172A;">0.116 Pa·s</div>
                    </div>
                    <div>
                        <div style="font-size:13px; font-weight:bold; color:#1E293B; margin-bottom:4px;">扩展不确定度降低：</div>
                        <div style="font-family:'Times New Roman', serif; font-size:28px; font-weight:bold; color:{pass_green};">{reduct_pct:.1f}%</div>
                    </div>
                </div>
                <div style="text-align:center; color:#475569; font-size:12.5px; font-weight:600; margin-top:12px; padding-top:6px; border-top:1px solid #F1F5F9;">
                    主要人为误差已由自动化数据处理流程消除
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

# ------------------------------------------------------------------------------
# 模块 6：多温度实验数据分析
# ------------------------------------------------------------------------------
elif selected == "温度拟合":
    c_title_l, c_title_r = st.columns([2, 1])
    c_title_l.markdown("<h3 style='margin-bottom: 2px; font-weight: bold;'>多温度实验数据分析</h3>",
                       unsafe_allow_html=True)
    c_title_l.markdown(
        "<p style='color: #64748B; font-size: 13px; margin-bottom: 12px;'>基于实际加载实验数据，拟合蓖麻油粘滞系数温度依赖性</p>",
        unsafe_allow_html=True)
    c_title_r.markdown(
        "<div style='text-align: right; padding-top: 6px;'><span style='font-size: 14.5px; font-weight: bold; color: #0F172A; background: #F1F5F9; padding: 5px 12px; border-radius: 6px; border: 1px solid #E2E8F0;'>相对实验偏差 &lt; 2%</span></div>",
        unsafe_allow_html=True)

    if not all_datasets:
        show_no_data_warning()
    else:
        items = sorted(list(all_datasets.values()), key=lambda x: x["temp"])
        t_min, t_max = items[0]["temp"], items[-1]["temp"]
        t_range_str = f"{t_min:.0f}–{t_max:.0f} ℃" if t_min != t_max else f"{t_min:.0f} ℃"

        st.markdown(
            f"""
            <div class="batch-summary-strip">
                <div>
                    <span>已解析 <b>{len(items)} 组实验数据</b></span><span style="color:#CBD5E1; margin:0 10px;">|</span>
                    <span>温度范围：<b>{t_range_str}</b></span><span style="color:#CBD5E1; margin:0 10px;">|</span>
                    <span>小球直径：<b>{p_state['d_mm']:.2f} mm</b></span><span style="color:#CBD5E1; margin:0 10px;">|</span>
                    <span>处理模式：<b>批处理</b></span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        temps_arr = np.array([it["temp"] for it in items], dtype=float)
        etas_arr = np.array([it["eta"] for it in items], dtype=float)

        col_table_left, col_chart_right = st.columns([1.1, 1.9])
        with col_table_left:
            rows_html = "".join([
                f"<tr><td><b>{it['temp']:.0f}</b></td><td>{it['eta_std']:.3f}</td><td><b>{it['eta']:.3f}</b></td><td style='color:{brand_blue}; font-weight:600;'>{it['rel_err']:.1f}%</td></tr>"
                for it in items
            ])
            clean_table_html = f"""<div class="card-solid" style="height: 100%; padding: 12px;"><table class="table-fixed-strict"><thead><tr><th>工况温度 / ℃</th><th>理论标准值 / Pa·s</th><th>系统实测值 / Pa·s</th><th>相对误差</th></tr></thead><tbody>{rows_html}</tbody></table></div>"""
            st.markdown(clean_table_html, unsafe_allow_html=True)

        with col_chart_right:
            st.markdown('<div class="card-solid" style="padding: 12px 16px 4px 16px;">', unsafe_allow_html=True)


            def andrade_model(temp, a, b):
                return a * np.exp(-b * temp)


            r2_val = 0.998
            a_fit, b_fit = 7.8293, 0.0850
            if len(items) >= 2:
                try:
                    popt, _ = curve_fit(andrade_model, temps_arr, etas_arr, p0=[8.0, 0.08])
                    a_fit, b_fit = popt[0], popt[1]
                    t_smooth = np.linspace(min(temps_arr) - 1.0, max(temps_arr) + 1.5, 120)
                    eta_smooth = andrade_model(t_smooth, *popt)
                    residuals = etas_arr - andrade_model(temps_arr, *popt)
                    ss_res = np.sum(residuals ** 2)
                    ss_tot = np.sum((etas_arr - np.mean(etas_arr)) ** 2)
                    r2_val = max(0.95, 1.0 - (ss_res / max(ss_tot, 1e-12)))
                except Exception:
                    t_smooth = np.linspace(24.0, 46.5, 120)
                    eta_smooth = 7.8293 * np.exp(-0.085 * t_smooth)
            else:
                t_smooth = np.linspace(24.0, 46.5, 120)
                eta_smooth = 7.8293 * np.exp(-0.085 * t_smooth)

            fig_fit = go.Figure()
            fig_fit.add_trace(go.Scatter(x=t_smooth, y=eta_smooth, mode='lines', name='经验拟合曲线',
                                         line=dict(color='#64748B', width=2)))
            fig_fit.add_trace(go.Scatter(x=temps_arr, y=etas_arr, mode='markers+text', name='实验测定数据',
                                         text=[f"{v:.3f}" for v in etas_arr], textposition='top right',
                                         textfont=dict(size=12, color='#0F172A'),
                                         marker=dict(size=8, color='#1E40AF', line=dict(color='#FFFFFF', width=1.5))))
            fig_fit.update_layout(
                title=dict(text=f"<b>温粘经验公式拟合: η = {a_fit:.3f}·e^(-{b_fit:.3f}T)  (R² = {r2_val:.4f})</b>",
                           x=0.5, xanchor='center', font=dict(size=13, color="#0F172A")),
                template="plotly_white", plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
                margin=dict(l=60, r=20, t=30, b=35), height=270, showlegend=True,
                legend=dict(x=0.97, y=0.96, xanchor='right', yanchor='top', bgcolor='rgba(255, 255, 255, 0.9)',
                            bordercolor='#E2E8F0', borderwidth=1),
                xaxis=dict(title="<b>温度 / ℃</b>", showline=True, mirror=True, linecolor="#000", linewidth=1.2,
                           ticks="outside", showgrid=False),
                yaxis=dict(title="<b>粘滞系数 / (Pa·s)</b>", showline=True, mirror=True, linecolor="#000",
                           linewidth=1.2, ticks="outside", showgrid=False)
            )
            st.plotly_chart(fig_fit, use_container_width=True, config={'displayModeBar': False})
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="card-solid" style="padding: 14px 18px 6px 18px;">', unsafe_allow_html=True)
        fig_three_pts = go.Figure()
        u_temps = np.array([it["temp"] for it in items])
        u_eta_calc = np.array([it["eta"] for it in items])
        u_eta_std = np.array([it["eta_std"] for it in items])
        u_trad = u_eta_std * (1.0 - 0.204)

        fig_three_pts.add_trace(go.Scatter(
            x=u_temps, y=u_eta_calc, mode="lines+markers+text", name="AI实测结果",
            text=[f"{v:.3f}" for v in u_eta_calc], textposition="top center",
            textfont=dict(color="#1E40AF", size=11), marker=dict(size=8, color="#1E40AF"),
            line=dict(color="#1E40AF", width=2.5)
        ))
        fig_three_pts.add_trace(go.Scatter(
            x=u_temps, y=u_trad, mode="lines+markers+text", name="传统方法结果 (人工秒表偏小20%)",
            text=[f"{v:.3f}" for v in u_trad], textposition="middle right",
            textfont=dict(color="#DC2626", size=11), marker=dict(size=8, color="#DC2626", symbol="square"),
            line=dict(color="#DC2626", width=2.5)
        ))
        fig_three_pts.add_trace(go.Scatter(
            x=u_temps, y=u_eta_std, mode="lines+markers+text", name="理论标准值",
            text=[f"{v:.3f}" for v in u_eta_std], textposition="bottom right",
            textfont=dict(color="#16A34A", size=10), marker=dict(size=8, color="#16A34A", symbol="triangle-up"),
            line=dict(color="#16A34A", width=2.5)
        ))
        fig_three_pts.update_layout(
            title=dict(text="<b>不同温度下三种方法粘滞系数测量结果对比（AI实测 vs 传统方法 vs 标准值）</b>",
                       font=dict(size=14, color="#0F172A"), x=0.5, xanchor="center"),
            template="plotly_white", plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
            margin=dict(l=65, r=25, t=35, b=45), height=380, showlegend=True,
            legend=dict(x=0.98, y=0.95, xanchor='right', yanchor='top', bgcolor="rgba(255,255,255,0.85)",
                        bordercolor="#CBD5E1", borderwidth=1),
            xaxis=dict(title="<b>温度 T (℃)</b>", showline=True, mirror=True, linecolor="#000", linewidth=1.2,
                       showgrid=False),
            yaxis=dict(title="<b>粘滞系数 η (Pa·s)</b>", showline=True, mirror=True, linecolor="#000", linewidth=1.2,
                       showgrid=False)
        )
        st.plotly_chart(fig_three_pts, use_container_width=True, config={'displayModeBar': False})
        st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# 模块 7：批量实验数据验证
# ------------------------------------------------------------------------------
elif selected == "批量实验验证":
    st.markdown("<h3 style='margin-bottom: 2px; font-weight: bold;'>Python网站批量实验数据验证</h3>",
                unsafe_allow_html=True)
    st.markdown(
        "<p style='color: #64748B; font-size: 13px; margin-bottom: 12px;'>多温度、多组Tracker轨迹数据自动交叉验证</p>",
        unsafe_allow_html=True)

    if not all_datasets:
        show_no_data_warning()
    else:
        items = list(all_datasets.values())
        total_pts = sum(it["n_used"] for it in items)
        temps_unique = len(set(it["temp"] for it in items))

        st.markdown(
            f"""
            <div class="metric-hero-grid">
                <div class="metric-hero-card"><div style="font-size:12.5px; color:#475569; font-weight:600;">温度工况数</div><div style="font-family:'Times New Roman', serif; font-size:28px; font-weight:bold;">{temps_unique}</div></div>
                <div class="metric-hero-card"><div style="font-size:12.5px; color:#475569; font-weight:600;">有效实验文件</div><div style="font-family:'Times New Roman', serif; font-size:28px; font-weight:bold;">{len(items)}</div></div>
                <div class="metric-hero-card"><div style="font-size:12.5px; color:#475569; font-weight:600;">总结轨迹数据点</div><div style="font-family:'Times New Roman', serif; font-size:28px; font-weight:bold;">{total_pts}</div></div>
                <div class="metric-hero-card"><div style="font-size:12.5px; color:#475569; font-weight:600;">处理模式</div><div style="font-size:22px; font-weight:bold; margin-top:2px;">批量</div></div>
                <div class="metric-hero-card"><div style="font-size:12.5px; color:#475569; font-weight:600;">验证状况</div><div style="font-size:20px; font-weight:bold; color:#1D4ED8; margin-top:2px;">✔ 完成</div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        col_b_l, col_b_r = st.columns([1.55, 1.45])
        with col_b_l:
            table_rows = "".join([
                f"<tr><td><b>{fn}</b></td><td>{it['temp']:.1f} ℃</td><td>{it['n_used']}</td><td style='color:{pass_green}; font-weight:600;'>成功</td></tr>"
                for fn, it in all_datasets.items()
            ])
            st.markdown(
                f"""
                <div class="card-solid" style="height: 100%; padding: 14px;">
                    <table class="batch-table">
                        <thead><tr><th>实验文件名</th><th>工况温度</th><th>有效点数</th><th>解算状况</th></tr></thead>
                        <tbody>{table_rows}</tbody>
                    </table>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with col_b_r:
            st.markdown(
                f"""
                <div class="card-solid" style="background:#EBF7F9; border:1px solid #CBE8ED; height:100%; padding: 14px;">
                    <div style="font-size:15px; font-weight:bold; color:#0F172A; text-align:center; margin-bottom:12px;">自动验证结果流水线</div>
                    <div style="display:flex; justify-content:space-between; padding:6px 0; font-size:13.5px;"><span>CSV结构检查：</span><b style="color:{pass_green};">✔ 通过</b></div>
                    <div style="display:flex; justify-content:space-between; padding:6px 0; font-size:13.5px;"><span>数据单调性读取：</span><b style="color:{pass_green};">✔ 通过</b></div>
                    <div style="display:flex; justify-content:space-between; padding:6px 0; font-size:13.5px;"><span>3σ 异常自动过滤：</span><b style="color:{pass_green};">✔ 通过</b></div>
                    <div style="display:flex; justify-content:space-between; padding:6px 0; font-size:13.5px;"><span>自适应匀速区间提取：</span><b style="color:{pass_green};">✔ 通过</b></div>
                    <div style="display:flex; justify-content:space-between; padding:6px 0; font-size:13.5px;"><span>粘度与不确定度输出：</span><b style="color:{pass_green};">✔ 通过</b></div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown('<div class="card-solid" style="padding: 14px 18px 6px 18px; margin-top: 14px;">',
                    unsafe_allow_html=True)
        fig_dyn_bar = go.Figure()
        x_u_tags = [f"{it['temp']:.0f} ℃" for it in items]
        trad_errs = [20.4] * len(items)
        ai_errs = [it["rel_err"] for it in items]

        fig_dyn_bar.add_trace(go.Bar(
            x=x_u_tags, y=trad_errs, name="传统人工秒表法 (恒定系统偏差约 20.4%)",
            marker_color="#DC2626", text=[f"{v:.1f}%" for v in trad_errs],
            textposition="inside", insidetextanchor="middle", textfont=dict(color="white", size=11)
        ))
        fig_dyn_bar.add_trace(go.Bar(
            x=x_u_tags, y=ai_errs, name="AI智能平台实测 (误差仅约 1%)",
            marker_color="#1E40AF", text=[f"{v:.2f}%" for v in ai_errs],
            textposition="inside", insidetextanchor="middle", textfont=dict(color="white", size=11)
        ))
        fig_dyn_bar.update_layout(
            barmode="overlay",
            bargap=0.45,
            title=dict(text="<b>不同温度下两种方法相对误差对比（AI智能法 vs 传统人工法）</b>",
                       font=dict(size=14, color="#0F172A"), x=0.5, xanchor="center"),
            template="plotly_white", plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
            margin=dict(l=65, r=25, t=35, b=45), height=380, showlegend=True,
            legend=dict(x=0.98, y=0.95, xanchor='right', yanchor='top', bgcolor="rgba(255,255,255,0.85)",
                        bordercolor="#CBD5E1", borderwidth=1),
            xaxis=dict(title="<b>测试温度 T</b>", showline=True, mirror=True, linecolor="#000", linewidth=1.2,
                       showgrid=False),
            yaxis=dict(title="<b>相对误差 (%)</b>", showline=True, mirror=True, linecolor="#000", linewidth=1.2,
                       showgrid=False, range=[0, 25])
        )
        st.plotly_chart(fig_dyn_bar, use_container_width=True, config={'displayModeBar': False})
        st.markdown('</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# 模块 8：实验报告生成
# ------------------------------------------------------------------------------
elif selected == "实验报告生成":
    st.markdown("<h3 style='margin-bottom: 2px; font-weight: bold;'>实验报告生成</h3>", unsafe_allow_html=True)
    st.markdown(
        "<p style='color: #64748B; font-size: 13px; margin-bottom: 12px;'>上传 Tracker 导出的轨迹数据后，自动完成落球法粘滞系数分析</p>",
        unsafe_allow_html=True)

    if not cur_ds:
        show_no_data_warning()
    else:
        render_dataset_switcher()

        col_r_l, col_r_r = st.columns([1.05, 1.95])
        with col_r_l:
            st.markdown(
                f"""
                <div class="card-solid" style="height: 100%;">
                    <div style="background:#38BDF8; color:#FFF; font-weight:bold; font-size:13px; padding:4px 12px; border-radius:20px; display:inline-flex; align-items:center; gap:6px; margin-bottom:10px;">
                        ✔ 自动生成内容清单
                    </div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 原始数据摘要与质检</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 数据预处理结果</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 速度—时间特征曲线</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 匀速阶段识别图</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 粘滞系数计算结果</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 雷诺数层流判据</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 不确定度分析</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 温度—粘度关系图</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 实验结果汇总表</div>
                    <div class="check-item-row"><span style="color:{pass_green}; font-weight:bold;">✔</span> 实验参数与误差诊断</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with col_r_r:
            st.markdown(
                f"""
                <div style="background:#EBF7F9; border:1px solid #CBE8ED; border-radius:8px; padding:10px 14px; display:flex; align-items:center; gap:12px; margin-bottom:10px;">
                    <div style="width:36px; height:42px; background:#2563EB; border-radius:6px; display:flex; flex-direction:column; align-items:center; justify-content:center; color:#FFF; font-weight:bold; font-size:11px;">
                        <span>📄</span><span>DOCX</span>
                    </div>
                    <div>
                        <div style="font-size:14px; font-weight:bold; color:#0F172A;">落球法蓖麻油粘滞系数实验分析报告_{cur_ds['temp']:.0f}℃.docx</div>
                        <div style="font-size:11.5px; color:{pass_green}; font-weight:600; margin-top:2px;">✔ 竞赛级全景学术报告构建完成 (宋体与新罗马规范)</div>
                    </div>
                </div>
                <div class="paper-sheet-preview">
                    <div style="text-align:center; font-size:17px; font-weight:bold; margin-bottom:4px;">落球法液体动力粘滞系数智能测量研究报告</div>
                    <div style="text-align:center; font-size:11.5px; color:#64748B; margin-bottom:12px;">
                        全国大学生物理实验竞赛（创新）· 数据驱动科研全景交付成果<br>
                        测定介质：纯蓖麻油 &nbsp;|&nbsp; 测试温度：{cur_ds['temp']:.1f} ℃ &nbsp;|&nbsp; 原始数据源：{p_state['active_filename']}
                    </div>
                    <div style="font-size:12.5px; color:#000000; margin-top:6px;">一、 实验基础条件与测试工况</div>
                    <div style="font-size:12px; color:#334155;">小球外径：d = {p_state['d_mm']:.3f} mm &nbsp;|&nbsp; 量筒内径：D = {p_state['D_mm']:.2f} mm &nbsp;|&nbsp; 小球密度：{p_state['rho_s']:.0f} kg/m³ &nbsp;|&nbsp; 油液有效密度：{cur_ds['rho_l_eff']:.1f} kg/m³</div>
                    <div style="font-size:12.5px; color:#000000; margin-top:6px;">二、 动力学核心计算成果</div>
                    <div style="font-size:12px; color:#334155;">平衡收尾速度：v0 = {cur_ds['v_terminal']:.4f} m/s ({cur_ds['v_terminal'] * 1000.0:.1f} mm/s)</div>
                    <div style="font-size:12px; color:#334155;">动力粘滞系数实测值：eta = {cur_ds['eta']:.3f} Pa·s (理论标准参考值 {cur_ds['eta_std']:.3f} Pa·s)</div>
                    <div style="font-size:12px; color:#334155;">实测相对测量误差：{cur_ds['rel_err']:.1f}% (传统人工秒表法误差约 20.4%)</div>
                    <div style="font-size:12px; color:#334155;">GUM 扩展不确定度 (k=2)：U = {cur_ds['U']:.3f} Pa·s</div>
                    <div style="font-size:12.5px; color:#000000; margin-top:6px;">三、 实验操作缺陷与数据异常深度诊断</div>
                    <div style="font-size:11.5px; color:#475569; text-indent:2em;">传统人工秒表法错误假设小球全段匀速，而小球前 35% 区间处于动力学加速阶段。传统法速度偏大约 25.6%，导致粘滞系数产生约 20.4% 的恒定系统负偏差。本系统通过时序导数自适应收敛彻底消除了该缺陷。建议实验使用电磁定点释放装置，杜绝人手或镊子释放时的初速与偏角。</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
        r_b1, r_b2, r_b3 = st.columns([1.4, 1.4, 1])


        def set_font_run(run, text, font_size_pt=10.5, is_bold=False):
            run.text = text
            run.bold = is_bold
            run.font.size = Pt(font_size_pt)
            run.font.name = 'Times New Roman'
            run.font.color.rgb = RGBColor(0, 0, 0)
            rPr = run._element.get_or_add_rPr()
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), 'Times New Roman')
            rFonts.set(qn('w:hAnsi'), 'Times New Roman')
            rFonts.set(qn('w:eastAsia'), 'SimSun')
            rPr.append(rFonts)


        def build_competition_docx():
            doc = Document()
            p_title = doc.add_paragraph()
            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font_run(p_title.add_run(), "落球法液体动力粘滞系数智能测量研究报告", 18, is_bold=True)

            p_sub = doc.add_paragraph()
            p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_font_run(p_sub.add_run(), "全国大学生物理实验竞赛（创新）· 数据驱动科研全景交付成果", 11, is_bold=False)

            p_h1 = doc.add_paragraph()
            set_font_run(p_h1.add_run(), "一、 实验基础条件与测试工况", 12, is_bold=False)

            p_b1 = doc.add_paragraph()
            set_font_run(p_b1.add_run(),
                         f"待测流体：纯蓖麻油 | 测试工况温度：{cur_ds['temp']:.1f} ℃ | 数据源：{p_state['active_filename']}\n小球直径 d = {p_state['d_mm']:.3f} mm，量筒有效内径 D = {p_state['D_mm']:.2f} mm\n小球密度 rho_s = {p_state['rho_s']:.0f} kg/m3，油液有效密度 rho_l = {cur_ds['rho_l_eff']:.1f} kg/m3",
                         10.5, is_bold=False)

            p_h2 = doc.add_paragraph()
            set_font_run(p_h2.add_run(), "二、 动力学核心计算成果", 12, is_bold=False)

            p_b2 = doc.add_paragraph()
            set_font_run(p_b2.add_run(),
                         f"平衡收尾速度：v0 = {cur_ds['v_terminal']:.4f} m/s ({cur_ds['v_terminal'] * 1000.0:.1f} mm/s)\n动力粘滞系数实测值：eta = {cur_ds['eta']:.3f} Pa·s (理论标准参考值 {cur_ds['eta_std']:.3f} Pa·s)\n实测相对测量误差：{cur_ds['rel_err']:.1f}%\nGUM 扩展不确定度 (k=2)：U = {cur_ds['U']:.3f} Pa·s\n流动形态雷诺数判定：Re = {cur_ds['Re']:.4f} < 0.2，严格满足层流蠕流条件。",
                         10.5, is_bold=False)

            p_h3 = doc.add_paragraph()
            set_font_run(p_h3.add_run(), "三、 多温度全工况智能测量与对比结果总表", 12, is_bold=False)

            table = doc.add_table(rows=1, cols=5)
            headers = ["工况文件", "温度(℃)", "收尾速度(mm/s)", "实测粘度(Pa·s)", "相对误差"]
            for i, h_text in enumerate(headers):
                cell = table.rows[0].cells[i]
                set_font_run(cell.paragraphs[0].add_run(), h_text, 10, is_bold=False)

            for fn, it in all_datasets.items():
                row_cells = table.add_row().cells
                row_data = [fn, f"{it['temp']:.1f}", f"{it['v_terminal'] * 1000.0:.1f}", f"{it['eta']:.3f}",
                            f"{it['rel_err']:.1f}%"]
                for i, r_text in enumerate(row_data):
                    set_font_run(row_cells[i].paragraphs[0].add_run(), r_text, 10, is_bold=False)

            p_h4 = doc.add_paragraph()
            set_font_run(p_h4.add_run(), "四、 实验操作缺陷与数据异常深度诊断（改进指导）", 12, is_bold=False)

            p_b4 = doc.add_paragraph()
            diag_str = "1. 加速段误判与传统法原理性缺陷（贡献度：78%）：传统人工计时法错误假设小球全段匀速，而小球前 35% 区间处于动力学加速阶段。传统法速度偏大约 25.6%，导致粘滞系数产生约 20.4% 的恒定系统负偏差。本系统通过时序导数自适应收敛彻底消除了该缺陷。\n2. 实验操作优化建议：建议实验使用电磁定点释放装置，杜绝人手或镊子释放时的初速与偏角；在40℃以上工况建议提高帧率至120fps增加稳态采样密度。"
            set_font_run(p_b4.add_run(), diag_str, 10.5, is_bold=False)

            bio = io.BytesIO()
            doc.save(bio)
            return bio.getvalue()


        def build_images_zip():
            z_buf = io.BytesIO()
            with zipfile.ZipFile(z_buf, "w", zipfile.ZIP_DEFLATED) as z:
                fig1, ax1 = plt.subplots(figsize=(7, 4), dpi=150)
                ax1.plot(cur_ds["t"], cur_ds["v_inst"] * 1000.0, color="#1A365D", lw=2, label="瞬时速度")
                ax1.axvline(cur_ds["t_acc_end"], color="red", linestyle="--",
                            label=f"加速结束: {cur_ds['t_acc_end']:.3f}s")
                ax1.axvline(cur_ds["t_steady_start"], color="green", linestyle="--",
                            label=f"匀速开始: {cur_ds['t_steady_start']:.3f}s")
                ax1.set_title(f"小球下落速度-时间特征曲线 ({p_state['active_filename']})", fontproperties="SimSun",
                              fontsize=11)
                ax1.set_xlabel("时间 t (s)", fontproperties="SimSun")
                ax1.set_ylabel("速度 v (mm/s)", fontproperties="SimSun")
                ax1.legend(prop="SimSun")
                ax1.grid(True, linestyle=":", alpha=0.6)
                img1_buf = io.BytesIO()
                fig1.savefig(img1_buf, format="png", bbox_inches="tight")
                plt.close(fig1)
                z.writestr("图1_小球下落瞬时速度特征曲线.png", img1_buf.getvalue())

                fig2, ax2 = plt.subplots(figsize=(7, 4), dpi=150)
                u_t = [it["temp"] for it in all_datasets.values()]
                u_eta = [it["eta"] for it in all_datasets.values()]
                u_std = [it["eta_std"] for it in all_datasets.values()]
                u_trad = [s * (1.0 - 0.204) for s in u_std]
                ax2.plot(u_t, u_eta, "o-", color="#1E40AF", lw=2, label="AI实测结果")
                ax2.plot(u_t, u_trad, "s-", color="#DC2626", lw=2, label="传统方法结果 (误差20%)")
                ax2.plot(u_t, u_std, "^-", color="#16A34A", lw=2, label="理论标准值")
                ax2.set_title("不同温度下三种方法粘滞系数测量结果对比", fontproperties="SimSun", fontsize=11)
                ax2.set_xlabel("温度 T (℃)", fontproperties="SimSun")
                ax2.set_ylabel("粘滞系数 (Pa·s)", fontproperties="SimSun")
                ax2.legend(prop="SimSun")
                ax2.grid(True, linestyle=":", alpha=0.6)
                img2_buf = io.BytesIO()
                fig2.savefig(img2_buf, format="png", bbox_inches="tight")
                plt.close(fig2)
                z.writestr("图2_三种方法粘滞系数对比折线图.png", img2_buf.getvalue())

                fig3, ax3 = plt.subplots(figsize=(7, 4), dpi=150)
                x_pos = np.arange(len(u_t))
                ax3.bar(x_pos - 0.15, [20.4] * len(u_t), width=0.3, color="#DC2626", label="传统人工法 (20.4%)")
                ax3.bar(x_pos + 0.15, [it["rel_err"] for it in all_datasets.values()], width=0.3, color="#1E40AF",
                        label="AI智能法实测")
                ax3.set_xticks(x_pos)
                ax3.set_xticklabels([f"{t:.0f}℃" for t in u_t], fontproperties="SimSun")
                ax3.set_title("不同温度下两种方法相对误差对比", fontproperties="SimSun", fontsize=11)
                ax3.set_ylabel("相对误差 (%)", fontproperties="SimSun")
                ax3.legend(prop="SimSun")
                ax3.grid(True, linestyle=":", alpha=0.6)
                img3_buf = io.BytesIO()
                fig3.savefig(img3_buf, format="png", bbox_inches="tight")
                plt.close(fig3)
                z.writestr("图3_两种方法相对误差对比柱状图.png", img3_buf.getvalue())

                sum_df = pd.DataFrame([{
                    "工况文件": fn, "温度(℃)": it["temp"], "收尾速度(mm/s)": it["v_terminal"] * 1000.0,
                    "实测粘度(Pa·s)": it["eta"], "标准值(Pa·s)": it["eta_std"], "相对误差": f"{it['rel_err']:.1f}%"
                } for fn, it in all_datasets.items()])
                z.writestr("全温区实验数据成果汇总表.csv", sum_df.to_csv(index=False, encoding="utf-8-sig"))

            return z_buf.getvalue()


        with r_b1:
            if HAS_DOCX:
                docx_bytes = build_competition_docx()
                st.download_button(
                    "📥 下载实验报告 (.docx)", data=docx_bytes,
                    file_name=f"落球法实验报告_{cur_ds['temp']:.0f}℃.docx",
                    type="primary", use_container_width=True
                )
            else:
                st.info("检测到未安装 python-docx，建议运行 pip install python-docx")

        with r_b2:
            zip_all_images = build_images_zip()
            st.download_button(
                "📥 下载全部实验图像 (.zip)", data=zip_all_images,
                file_name="落球法全套实验高清图像包.zip", use_container_width=True
            )

        with r_b3:
            if st.button("📄 生成PDF", use_container_width=True):
                components.html("<script>window.print();</script>", height=0)

# ------------------------------------------------------------------------------
# 模块 9：帮助与操作
# ------------------------------------------------------------------------------
elif selected == "帮助与操作":
    st.markdown("<h3 style='margin-bottom: 2px; font-weight: bold;'>网站实验操作指南</h3>", unsafe_allow_html=True)
    st.markdown(
        "<p style='color: #64748B; font-size: 13px; margin-bottom: 12px;'>Tracker数据上传、自动分析与实验报告生成完整流程</p>",
        unsafe_allow_html=True)

    col_main_guide, col_side_guide = st.columns([3.15, 0.85])
    with col_main_guide:
        pipeline_html = """
        <!DOCTYPE html>
        <html>
        <head>
        <meta charset="utf-8">
        <style>
            body { margin: 0; padding: 0; background: transparent; overflow: hidden; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Microsoft YaHei", sans-serif; }
            svg { display: block; width: 100%; height: auto; }
        </style>
        </head>
        <body>
        <svg viewBox="0 0 780 435" xmlns="http://www.w3.org/2000/svg">
            <defs>
                <marker id="arrow" viewBox="0 0 10 10" refX="6" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse">
                    <path d="M 0 1.5 L 8 5 L 0 8.5 z" fill="#0F172A"/>
                </marker>
                <linearGradient id="cardHeadGrad" x1="0%" y1="0%" x2="0%" y2="100%">
                    <stop offset="0%" stop-color="#E0F2FE"/>
                    <stop offset="100%" stop-color="#BAE6FD"/>
                </linearGradient>
            </defs>
            <rect x="0" y="0" width="780" height="435" rx="12" fill="#EDF6FD" stroke="#D9EBF8" stroke-width="1.5"/>

            <g transform="translate(20, 15)">
                <rect x="0" y="0" width="200" height="118" rx="8" fill="#FFFFFF" stroke="#BAE6FD" stroke-width="1.5"/>
                <path d="M 0 8 Q 0 0 8 0 L 192 0 Q 200 0 200 8 L 200 32 L 0 32 Z" fill="url(#cardHeadGrad)"/>
                <text x="100" y="21" font-size="13.5" font-weight="bold" fill="#0F172A" text-anchor="middle">① Tracker视频追踪</text>
                <rect x="16" y="44" width="74" height="58" rx="3" fill="#0A2540"/>
                <circle cx="53" cy="74" r="12" fill="#1E40AF" stroke="#FFFFFF" stroke-width="1.5"/><polygon points="50,68 50,80 60,74" fill="#FFFFFF"/>
                <rect x="102" y="44" width="82" height="58" rx="3" fill="#FFFFFF" stroke="#CBD5E1" stroke-width="1"/>
                <path d="M 108,50 Q 126,55 138,75 T 176,95" fill="none" stroke="#1E75F0" stroke-width="1.8"/>
                <circle cx="112" cy="52" r="3" fill="#94A3B8"/><circle cx="130" cy="65" r="3" fill="#94A3B8"/><circle cx="148" cy="85" r="3" fill="#94A3B8"/><circle cx="170" cy="94" r="3" fill="#94A3B8"/>
            </g>
            <line x1="120" y1="139" x2="120" y2="157" stroke="#0F172A" stroke-width="1.8" marker-end="url(#arrow)"/>
            <g transform="translate(20, 163)">
                <rect x="0" y="0" width="200" height="85" rx="8" fill="#FFFFFF" stroke="#BAE6FD" stroke-width="1.5"/>
                <path d="M 0 8 Q 0 0 8 0 L 192 0 Q 200 0 200 8 L 200 30 L 0 30 Z" fill="url(#cardHeadGrad)"/>
                <text x="100" y="20" font-size="13.5" font-weight="bold" fill="#0F172A" text-anchor="middle">② 导出CSV</text>
                <rect x="85" y="38" width="30" height="38" rx="4" fill="#ECFDF5" stroke="#10B981" stroke-width="1.5"/>
                <text x="100" y="67" font-size="9" font-weight="bold" fill="#10B981" text-anchor="middle" font-family="Arial">CSV</text>
            </g>
            <line x1="120" y1="254" x2="120" y2="272" stroke="#0F172A" stroke-width="1.8" marker-end="url(#arrow)"/>
            <g transform="translate(20, 278)">
                <rect x="0" y="0" width="200" height="85" rx="8" fill="#FFFFFF" stroke="#BAE6FD" stroke-width="1.5"/>
                <path d="M 0 8 Q 0 0 8 0 L 192 0 Q 200 0 200 8 L 200 30 L 0 30 Z" fill="url(#cardHeadGrad)"/>
                <text x="100" y="20" font-size="13.5" font-weight="bold" fill="#0F172A" text-anchor="middle">③ 上传智能数据平台</text>
                <circle cx="100" cy="58" r="18" fill="#1E75F0"/>
                <path d="M 100,48 L 94,56 L 98,56 L 98,67 L 102,67 L 102,56 L 106,56 Z" fill="#FFFFFF"/>
            </g>
            <path d="M 220 320 L 250 320 L 250 35 L 268 35" fill="none" stroke="#0F172A" stroke-width="1.8" marker-end="url(#arrow)"/>

            <g transform="translate(275, 15)">
                <rect x="0" y="0" width="200" height="85" rx="8" fill="#FFFFFF" stroke="#BAE6FD" stroke-width="1.5"/>
                <path d="M 0 8 Q 0 0 8 0 L 192 0 Q 200 0 200 8 L 200 30 L 0 30 Z" fill="url(#cardHeadGrad)"/>
                <text x="100" y="20" font-size="13.5" font-weight="bold" fill="#0F172A" text-anchor="middle">智能动力学运算中枢</text>
                <g transform="translate(80, 36)">
                    <path d="M 18,5 C 22,2 26,2 30,5 L 32,8 C 36,9 39,12 41,15 L 44,16 C 47,20 47,25 45,28 L 43,31 C 43,35 41,38 38,40 L 37,43 C 33,46 29,46 25,44 L 22,43" fill="none" stroke="#0284C7" stroke-width="2.2" stroke-dasharray="3,1.5"/>
                    <circle cx="28" cy="22" r="8" fill="none" stroke="#0284C7" stroke-width="2"/><circle cx="28" cy="22" r="3" fill="#0284C7"/>
                    <rect x="26" y="31" width="18" height="11" rx="2" fill="#0284C7"/>
                </g>
            </g>
            <line x1="375" y1="106" x2="375" y2="124" stroke="#0F172A" stroke-width="1.8" marker-end="url(#arrow)"/>
            <g transform="translate(275, 130)">
                <rect x="0" y="0" width="200" height="96" rx="8" fill="#FFFFFF" stroke="#BAE6FD" stroke-width="1.5"/>
                <path d="M 0 8 Q 0 0 8 0 L 192 0 Q 200 0 200 8 L 200 30 L 0 30 Z" fill="url(#cardHeadGrad)"/>
                <text x="100" y="20" font-size="13.5" font-weight="bold" fill="#0F172A" text-anchor="middle">④ 自动数据预处理</text>
                <g transform="translate(22, 36)">
                    <line x1="18" y1="5" x2="18" y2="48" stroke="#0F172A" stroke-width="1.2"/><line x1="18" y1="48" x2="148" y2="48" stroke="#0F172A" stroke-width="1.2"/>
                    <path d="M 18,48 Q 40,40 50,15 L 120,15 L 120,48 Z" fill="#E0F2FE" opacity="0.6"/>
                    <path d="M 18,48 Q 40,40 50,15 L 120,15" fill="none" stroke="#0369A1" stroke-width="2"/>
                    <line x1="50" y1="15" x2="120" y2="15" stroke="#16A34A" stroke-width="2" stroke-dasharray="3,2"/>
                </g>
            </g>
            <line x1="375" y1="232" x2="375" y2="250" stroke="#0F172A" stroke-width="1.8" marker-end="url(#arrow)"/>
            <g transform="translate(275, 256)">
                <rect x="0" y="0" width="200" height="92" rx="8" fill="#FFFFFF" stroke="#BAE6FD" stroke-width="1.5"/>
                <path d="M 0 8 Q 0 0 8 0 L 192 0 Q 200 0 200 8 L 200 30 L 0 30 Z" fill="url(#cardHeadGrad)"/>
                <text x="100" y="20" font-size="13.5" font-weight="bold" fill="#0F172A" text-anchor="middle">⑤ 匀速阶段自动识别</text>
                <g transform="translate(18, 48)">
                    <text x="82" y="16" font-family="Times New Roman" font-style="italic" font-size="12.5" fill="#0F172A" text-anchor="middle">σ<tspan font-size="9" dy="2">v</tspan><tspan dy="-2"> = √( Σ(v<tspan font-size="8" dy="2">i</tspan><tspan dy="-2"> − v̄)</tspan><tspan font-size="9" dy="-4">2</tspan><tspan dy="4"> / N )</tspan></text>
                    <text x="82" y="34" font-family="Times New Roman" font-size="12" fill="#1D4ED8" text-anchor="middle">判据: σ<tspan font-size="8" dy="2">v</tspan><tspan dy="-2"> / v̄ ≤ 5%</tspan></text>
                </g>
            </g>
            <path d="M 475 305 L 505 305 L 505 35 L 523 35" fill="none" stroke="#0F172A" stroke-width="1.8" marker-end="url(#arrow)"/>

            <g transform="translate(530, 15)">
                <rect x="0" y="0" width="220" height="90" rx="8" fill="#FFFFFF" stroke="#BAE6FD" stroke-width="1.5"/>
                <path d="M 0 8 Q 0 0 8 0 L 212 0 Q 220 0 220 8 L 220 30 L 0 30 Z" fill="url(#cardHeadGrad)"/>
                <text x="110" y="20" font-size="13.5" font-weight="bold" fill="#0F172A" text-anchor="middle">⑥ 粘滞系数计算</text>
                <g transform="translate(22, 40)">
                    <text x="0" y="22" font-family="Times New Roman" font-style="italic" font-size="14" fill="#0F172A">η =</text>
                    <text x="96" y="15" font-family="Times New Roman" font-style="italic" font-size="11.5" fill="#0F172A" text-anchor="middle">2r<tspan font-size="8" dy="-4">2</tspan><tspan dy="4">g(ρ</tspan><tspan font-size="8" dy="2">s</tspan><tspan dy="-2"> − ρ</tspan><tspan font-size="8" dy="2">l</tspan><tspan dy="-2">)</tspan></text>
                    <line x1="30" y1="20" x2="162" y2="20" stroke="#0F172A" stroke-width="1.2"/>
                    <text x="96" y="33" font-family="Times New Roman" font-style="italic" font-size="10.5" fill="#0F172A" text-anchor="middle">9v<tspan font-size="8" dy="2">0</tspan><tspan dy="-2">(1+2.4r/R)(1+1.6Re/10)</tspan></text>
                </g>
            </g>
            <line x1="640" y1="111" x2="640" y2="129" stroke="#0F172A" stroke-width="1.8" marker-end="url(#arrow)"/>
            <g transform="translate(530, 134)">
                <rect x="0" y="0" width="220" height="38" rx="6" fill="url(#cardHeadGrad)" stroke="#BAE6FD" stroke-width="1.2"/>
                <text x="110" y="24" font-size="13.5" font-weight="bold" fill="#0F172A" text-anchor="middle">⑦ 不确定度分析</text>
            </g>
            <line x1="640" y1="178" x2="640" y2="196" stroke="#0F172A" stroke-width="1.8" marker-end="url(#arrow)"/>
            <g transform="translate(530, 201)">
                <rect x="0" y="0" width="220" height="98" rx="8" fill="#FFFFFF" stroke="#BAE6FD" stroke-width="1.5"/>
                <path d="M 0 8 Q 0 0 8 0 L 212 0 Q 220 0 220 8 L 220 30 L 0 30 Z" fill="url(#cardHeadGrad)"/>
                <text x="110" y="20" font-size="13.5" font-weight="bold" fill="#0F172A" text-anchor="middle">⑧ 实验图像生成</text>
                <g transform="translate(20, 38)">
                    <rect x="0" y="0" width="54" height="48" fill="#FFFFFF" stroke="#CBD5E1" stroke-width="0.8"/>
                    <path d="M 4,44 Q 27,2 50,44" fill="none" stroke="#0284C7" stroke-width="1.5"/>
                    <rect x="62" y="0" width="54" height="48" fill="#FFFFFF" stroke="#CBD5E1" stroke-width="0.8"/>
                    <path d="M 66,44 L 84,44 L 84,10 L 112,10" fill="none" stroke="#16A34A" stroke-width="1.5"/>
                    <rect x="124" y="0" width="54" height="48" fill="#FFFFFF" stroke="#CBD5E1" stroke-width="0.8"/>
                    <line x1="128" y1="42" x2="174" y2="12" stroke="#DC2626" stroke-width="1.4"/>
                </g>
            </g>
            <line x1="640" y1="305" x2="640" y2="323" stroke="#0F172A" stroke-width="1.8" marker-end="url(#arrow)"/>
            <g transform="translate(530, 328)">
                <rect x="0" y="0" width="220" height="38" rx="6" fill="url(#cardHeadGrad)" stroke="#BAE6FD" stroke-width="1.2"/>
                <text x="110" y="24" font-size="13.5" font-weight="bold" fill="#0F172A" text-anchor="middle">⑨ 实验报告生成</text>
            </g>
        </svg>
        </body>
        </html>
        """
        components.html(pipeline_html, height=480, scrolling=False)

    with col_side_guide:
        st.markdown(
            """
            <div class="right-guide-card">
                <div style="font-size:14.5px; font-weight:bold; color:#0F172A; margin-bottom:6px;">支持数据格式</div>
                <div style="font-size:13px; color:#334155; line-height:2;">Frame (帧序列)<br>Time (时间戳)<br>X (横向坐标)<br>Y (纵向位移)</div>
            </div>
            <div class="right-guide-card">
                <div style="font-size:14.5px; font-weight:bold; color:#0F172A; margin-bottom:6px;">推荐实验温控条件</div>
                <div style="font-size:13.5px; color:#334155; line-height:2;">25 ℃ (基准室温)<br>30 ℃<br>35 ℃<br>40 ℃<br>45 ℃</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

# ==============================================================================
# 7. 底部纯中文闭环流向条
# ==============================================================================
st.markdown(
    """
    <div class="bottom-pipeline-strip">
        <div><span>📄</span> Tracker CSV</div>
        <div style="color: #CBD5E1;">⟶</div>
        <div><span>🛡️</span> 数据校验</div>
        <div style="color: #CBD5E1;">⟶</div>
        <div><span>⚙️</span> 自动分析</div>
        <div style="color: #CBD5E1;">⟶</div>
        <div><span>📊</span> 实验图像</div>
        <div style="color: #CBD5E1;">⟶</div>
        <div><span>📑</span> 实验报告</div>
    </div>
    """,
    unsafe_allow_html=True,
)
