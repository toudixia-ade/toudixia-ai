import hmac
import io
import os
import sqlite3
import hashlib
import secrets
from datetime import datetime

import pandas as pd
import streamlit as st

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

try:
    from docx import Document
except ImportError:
    Document = None


st.set_page_config(
    page_title="投递侠AI",
    page_icon="🦸",
    layout="wide",
    initial_sidebar_state="expanded"
)

DATA_FILE = "applications.csv"
USER_DB_FILE = "users.db"

ADMIN_USERNAME = "admin_ljp_2026"
ADMIN_REGISTER_CODE = "LJP_ADMIN_2026_ONLY_ME"

FREE_DAILY_ANALYSIS_LIMIT = 3

# 公测阶段：True = 普通用户不限次数，只记录使用次数
# 后面正式收费时，把它改成 False，免费用户就会每天限制 3 次
PUBLIC_BETA_UNLIMITED = True

MEMBERSHIP_TYPES = ["免费试用", "周卡", "月卡", "年卡"]
UNLIMITED_MEMBERSHIPS = ["周卡", "月卡", "年卡"]

PLAN_PRICES = {
    "周卡": "¥9.9",
    "月卡": "¥29.9",
    "年卡": "¥99",
}

PLAN_DURATIONS = {
    "周卡": "7天不限次数",
    "月卡": "30天不限次数",
    "年卡": "365天不限次数",
}

ALL_COLUMNS = [
    "用户名",
    "保存时间",
    "求职者身份",
    "教育背景",
    "经历概述",
    "主要求职方向",
    "公司名称",
    "岗位名称",
    "城市",
    "薪资",
    "岗位链接",
    "匹配度评分",
    "推荐结论",
    "推荐简历版本",
    "简历匹配详情",
    "岗位画像",
    "匹配优势",
    "风险点",
    "简历优化建议",
    "投递状态",
    "备注",
    "打招呼语",
    "岗位JD",
]

CATEGORY_KEYWORDS = {
    "人力资源 / HR": [
        "人力", "招聘", "HR", "hr", "人事", "候选人", "面试", "邀约",
        "薪酬", "绩效", "培训", "员工关系", "劳动关系", "入离职", "社保"
    ],
    "运营 / 内容 / 新媒体": [
        "运营", "内容", "新媒体", "小红书", "抖音", "公众号", "视频号",
        "社群", "用户", "活动", "增长", "转化", "留存", "互动", "选题",
        "文案", "短视频", "直播"
    ],
    "市场 / 品牌 / 广告": [
        "市场", "品牌", "广告", "营销", "传播", "推广", "策划",
        "公关", "媒介", "投放", "竞品", "消费者", "调研", "曝光"
    ],
    "行政 / 助理 / 文职": [
        "行政", "助理", "文员", "办公室", "会议", "接待", "档案",
        "文件", "后勤", "办公用品", "固定资产", "报销", "考勤"
    ],
    "数据 / 分析 / 工具": [
        "数据", "分析", "Excel", "表格", "SQL", "Python", "BI",
        "指标", "复盘", "统计", "报表", "看板", "转化率", "留存率"
    ],
    "销售 / 客户 / 商务": [
        "销售", "客户", "商务", "BD", "线索", "成交", "回款",
        "客情", "维护客户", "开拓客户", "业绩", "提成"
    ],
    "产品 / 项目": [
        "产品", "项目", "需求", "原型", "PRD", "流程", "推进",
        "协调", "交付", "测试", "上线", "产品经理"
    ],
    "设计 / 视频 / 创意": [
        "设计", "视觉", "海报", "图片", "剪辑", "视频", "脚本",
        "拍摄", "创意", "审美", "PS", "AI", "PR", "剪映"
    ],
}

RISK_KEYWORDS = [
    "3年", "三年", "5年", "五年", "经验丰富", "独立负责",
    "抗压", "高强度", "加班", "销售", "业绩", "提成",
    "出差", "地推", "陌拜", "电话销售"
]

FRESH_GRAD_KEYWORDS = [
    "应届", "毕业生", "校招", "管培", "实习", "助理", "可培养"
]


def inject_custom_css():
    st.markdown(
        """
        <style>
        :root {
            --primary: #2563eb;
            --primary-dark: #1d4ed8;
            --purple: #7c3aed;
            --bg: #f6f8fc;
            --card: #ffffff;
            --text: #0f172a;
            --muted: #64748b;
            --border: #e2e8f0;
        }

        html,
        body,
        .stApp {
            color-scheme: light !important;
        }

        html, body, [class*="css"] {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", "PingFang SC", sans-serif;
        }

        .stApp {
            background:
                radial-gradient(circle at 15% 25%, rgba(37, 99, 235, 0.22), transparent 34%),
                radial-gradient(circle at 85% 30%, rgba(124, 58, 237, 0.22), transparent 36%),
                radial-gradient(circle at 50% 85%, rgba(14, 165, 233, 0.16), transparent 34%),
                linear-gradient(-45deg, #f8fbff, #eef4ff, #f4f0ff, #eff6ff, #f8fbff);
            background-size:
                150% 150%,
                160% 160%,
                140% 140%,
                400% 400%;
            background-position:
                0% 0%,
                100% 0%,
                50% 100%,
                0% 50%;
            animation: safeBackgroundMove 8s ease-in-out infinite alternate;
            overflow-x: hidden;
        }

        @keyframes safeBackgroundMove {
            0% {
                background-position:
                    0% 0%,
                    100% 0%,
                    50% 100%,
                    0% 50%;
            }
            100% {
                background-position:
                    18% 14%,
                    82% 22%,
                    60% 82%,
                    100% 50%;
            }
        }

        [data-testid="stAppViewContainer"] {
            color: #0f172a !important;
        }

        .block-container {
            max-width: 1520px;
            padding-top: 2rem;
            padding-bottom: 3rem;
            padding-left: 3rem;
            padding-right: 3rem;
            position: relative;
            z-index: 1;
        }

        section[data-testid="stSidebar"] {
            position: relative;
            z-index: 2;
        }

        h1, h2, h3 {
            color: var(--text);
            letter-spacing: -0.02em;
        }

        [data-testid="stWidgetLabel"] p {
            color: #334155 !important;
            font-weight: 700 !important;
        }

        .auth-brand-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 28px;
        }

        .auth-logo {
            display: inline-flex;
            align-items: center;
            gap: 10px;
            font-size: 22px;
            font-weight: 900;
            color: #172554;
        }

        .auth-logo-icon {
            width: 38px;
            height: 38px;
            border-radius: 14px;
            background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
            color: white;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
            box-shadow: 0 12px 28px rgba(37, 99, 235, 0.28);
        }

        .auth-version {
            padding: 8px 14px;
            border-radius: 999px;
            background: rgba(255,255,255,0.72);
            border: 1px solid rgba(226,232,240,0.9);
            color: #475569;
            font-size: 13px;
            font-weight: 700;
        }

        .landing-hero {
            padding: 54px 56px;
            border-radius: 34px;
            background:
                linear-gradient(135deg, rgba(37,99,235,0.96) 0%, rgba(79,70,229,0.96) 48%, rgba(124,58,237,0.96) 100%);
            color: white;
            box-shadow: 0 28px 80px rgba(37, 99, 235, 0.28);
            min-height: 455px;
            position: relative;
            overflow: hidden;
        }

        .landing-hero:after {
            content: "";
            position: absolute;
            right: -120px;
            top: -120px;
            width: 360px;
            height: 360px;
            border-radius: 999px;
            background: rgba(255,255,255,0.14);
        }

        .landing-kicker {
            display: inline-block;
            padding: 8px 14px;
            border-radius: 999px;
            background: rgba(255,255,255,0.16);
            border: 1px solid rgba(255,255,255,0.24);
            font-size: 14px;
            font-weight: 750;
            margin-bottom: 22px;
        }

        .landing-title {
            font-size: 58px;
            line-height: 1.08;
            font-weight: 950;
            letter-spacing: -0.06em;
            margin-bottom: 22px;
            max-width: 760px;
            position: relative;
            z-index: 1;
        }

        .landing-subtitle {
            font-size: 19px;
            line-height: 1.9;
            opacity: 0.94;
            max-width: 760px;
            margin-bottom: 30px;
            position: relative;
            z-index: 1;
        }

        .landing-tags {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            max-width: 760px;
            position: relative;
            z-index: 1;
        }

        .landing-tag {
            padding: 9px 16px;
            border-radius: 999px;
            background: rgba(255,255,255,0.22);
            border: 1px solid rgba(255,255,255,0.38);
            color: #ffffff;
            font-size: 14px;
            font-weight: 850;
            box-shadow: 0 8px 18px rgba(37, 99, 235, 0.18);
        }

        .login-card-title {
            font-size: 28px;
            font-weight: 920;
            color: #0f172a;
            margin-bottom: 8px;
        }

        .login-card-subtitle {
            font-size: 14px;
            color: #64748b;
            line-height: 1.8;
            margin-bottom: 20px;
        }

        .feature-grid-title {
            font-size: 22px;
            font-weight: 930;
            margin-top: 32px;
            margin-bottom: 16px;
            color: #0f172a;
        }

        .value-card {
            padding: 24px 24px;
            border-radius: 24px;
            background: rgba(255,255,255,0.88);
            border: 1px solid rgba(226,232,240,0.95);
            box-shadow: 0 14px 42px rgba(15, 23, 42, 0.06);
            height: 100%;
        }

        .value-title {
            font-size: 18px;
            font-weight: 900;
            color: #0f172a;
            margin-bottom: 10px;
        }

        .value-text {
            font-size: 14px;
            color: #475569;
            line-height: 1.85;
        }

        .scenario-list {
            padding: 26px 30px;
            border-radius: 26px;
            background: rgba(255,255,255,0.88);
            border: 1px solid rgba(226,232,240,0.95);
            box-shadow: 0 14px 42px rgba(15, 23, 42, 0.06);
            margin-top: 28px;
        }

        .scenario-title {
            font-size: 22px;
            font-weight: 930;
            color: #0f172a;
            margin-bottom: 16px;
        }

        .scenario-item {
            display: flex;
            gap: 12px;
            align-items: flex-start;
            color: #334155;
            font-size: 15px;
            line-height: 1.8;
            margin-bottom: 10px;
        }

        .scenario-dot {
            width: 9px;
            height: 9px;
            border-radius: 99px;
            background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
            margin-top: 10px;
            flex-shrink: 0;
        }

        .hero-card {
            padding: 34px;
            border-radius: 28px;
            background:
                linear-gradient(135deg, rgba(255,255,255,0.95) 0%, rgba(239,246,255,0.95) 55%, rgba(245,243,255,0.95) 100%);
            border: 1px solid rgba(226, 232, 240, 0.9);
            box-shadow: 0 18px 55px rgba(15, 23, 42, 0.08);
            margin-bottom: 26px;
            min-height: 210px;
        }

        .hero-title {
            font-size: 42px;
            font-weight: 900;
            color: #0f172a;
            margin-bottom: 10px;
            letter-spacing: -0.04em;
        }

        .hero-subtitle {
            font-size: 17px;
            color: #475569;
            line-height: 1.9;
            max-width: 880px;
        }

        .tag {
            display: inline-block;
            padding: 7px 14px;
            border-radius: 999px;
            background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
            color: #ffffff;
            font-size: 13px;
            font-weight: 800;
            margin-right: 10px;
            margin-top: 14px;
            box-shadow: 0 8px 18px rgba(37, 99, 235, 0.22);
        }

        .member-card {
            padding: 26px 24px;
            border-radius: 28px;
            background: rgba(255, 255, 255, 0.94);
            border: 1px solid rgba(226, 232, 240, 0.95);
            box-shadow: 0 18px 55px rgba(15, 23, 42, 0.08);
            min-height: 210px;
            margin-bottom: 26px;
        }

        .member-card-title {
            font-size: 20px;
            font-weight: 920;
            color: #0f172a;
            margin-bottom: 10px;
        }

        .member-card-text {
            font-size: 14px;
            color: #64748b;
            line-height: 1.8;
            margin-bottom: 18px;
        }

        .member-status-pill {
            display: inline-block;
            padding: 7px 12px;
            border-radius: 999px;
            background: #eff6ff;
            color: #1d4ed8;
            font-size: 13px;
            font-weight: 850;
            margin-bottom: 14px;
        }

               /* 弹窗外层遮罩：覆盖整个网页 */
        div[data-baseweb="modal"] {
            position: fixed !important;
            inset: 0 !important;
            width: 100vw !important;
            height: 100vh !important;
            max-width: none !important;
            max-height: none !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            background: rgba(15, 23, 42, 0.56) !important;
            z-index: 999999 !important;
        }

        /* 弹窗内容本体：居中且加宽 */
        div[data-baseweb="modal"] div[role="dialog"] {
            width: min(1180px, 92vw) !important;
            max-width: 1180px !important;
            margin: 0 auto !important;
            border-radius: 24px !important;
            overflow: hidden !important;
        }

        /* 兼容 Streamlit dialog 容器 */
        div[data-testid="stDialog"] {
            width: 100% !important;
            max-width: none !important;
        }

        div[data-testid="stDialog"] > div {
            margin-left: auto !important;
            margin-right: auto !important;
        }

        .pricing-header {
            text-align: center;
            margin-bottom: 22px;
        }

        .pricing-title {
            font-size: 30px;
            font-weight: 950;
            color: #0f172a;
            margin-bottom: 8px;
        }

        .pricing-subtitle {
            font-size: 14px;
            color: #64748b;
            line-height: 1.8;
        }

        .plan-card {
            padding: 26px 28px;
            border-radius: 24px;
            background: #ffffff;
            border: 1px solid #e2e8f0;
            box-shadow: 0 12px 34px rgba(15, 23, 42, 0.06);
            min-height: 320px;
        }

        .plan-card-featured {
            background: linear-gradient(180deg, #eff6ff 0%, #ffffff 100%);
            border: 1px solid #bfdbfe;
            box-shadow: 0 18px 46px rgba(37, 99, 235, 0.15);
        }

        .plan-name {
            font-size: 24px;
            font-weight: 930;
            color: #0f172a;
            margin-bottom: 10px;
            white-space: nowrap;
        }

        .plan-price {
            font-size: 42px;
            font-weight: 950;
            color: #111827;
            margin-bottom: 8px;
            white-space: nowrap;
            line-height: 1.1;
        }

        .plan-duration {
            font-size: 14px;
            color: #64748b;
            margin-bottom: 18px;
            white-space: nowrap;
        }

        .plan-feature {
            font-size: 14px;
            color: #334155;
            line-height: 1.9;
            margin-bottom: 6px;
        }

        .plan-badge {
            display: inline-block;
            padding: 6px 10px;
            border-radius: 999px;
            background: #dbeafe;
            color: #1d4ed8;
            font-size: 12px;
            font-weight: 850;
            margin-bottom: 12px;
        }

        .pay-box {
            padding: 18px 20px;
            border-radius: 18px;
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            color: #334155;
            font-size: 14px;
            line-height: 1.9;
            margin-top: 18px;
        }

        .section-title {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 22px;
            font-weight: 850;
            color: #0f172a;
            margin-top: 14px;
            margin-bottom: 12px;
        }

        .section-subtitle {
            font-size: 14px;
            color: var(--muted);
            margin-top: -4px;
            margin-bottom: 16px;
        }

        .metric-card {
            padding: 22px 18px;
            border-radius: 22px;
            background: rgba(255, 255, 255, 0.92);
            border: 1px solid rgba(226, 232, 240, 0.95);
            box-shadow: 0 10px 28px rgba(15, 23, 42, 0.06);
            min-height: 132px;
        }

        .metric-label {
            font-size: 14px;
            color: #64748b;
            margin-bottom: 8px;
            font-weight: 650;
        }

        .metric-value {
            font-size: 28px;
            color: #2563eb;
            font-weight: 900;
            line-height: 1.25;
            word-break: break-word;
        }

        .metric-note {
            font-size: 13px;
            color: #64748b;
            margin-top: 8px;
        }

        .info-card {
            padding: 22px 24px;
            border-radius: 22px;
            background: rgba(255, 255, 255, 0.94);
            border: 1px solid rgba(226, 232, 240, 0.95);
            box-shadow: 0 10px 30px rgba(15, 23, 42, 0.045);
            margin-bottom: 16px;
        }

        .info-card-title {
            font-size: 18px;
            font-weight: 850;
            color: #111827;
            margin-bottom: 10px;
        }

        .info-card-text {
            font-size: 15px;
            color: #374151;
            line-height: 1.85;
        }

        .status-pill {
            display: inline-block;
            padding: 7px 12px;
            border-radius: 999px;
            font-size: 13px;
            font-weight: 800;
            margin-bottom: 10px;
        }

        .status-success {
            background: #ecfdf5;
            color: #047857;
        }

        .status-warning {
            background: #fff7ed;
            color: #c2410c;
        }

        .status-danger {
            background: #fef2f2;
            color: #b91c1c;
        }

        .user-badge {
            padding: 14px 16px;
            border-radius: 18px;
            background: linear-gradient(135deg, #eff6ff 0%, #f5f3ff 100%);
            border: 1px solid #dbeafe;
            color: #1e3a8a;
            font-weight: 800;
            margin-bottom: 14px;
        }

        div.stButton > button {
            border-radius: 14px;
            border: none;
            background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
            color: white;
            font-weight: 800;
            min-height: 44px;
            box-shadow: 0 10px 22px rgba(37, 99, 235, 0.20);
            transition: all 0.2s ease;
        }

        div.stButton > button:hover {
            transform: translateY(-1px);
            filter: brightness(0.98);
            color: white;
            border: none;
        }

        div.stDownloadButton > button {
            border-radius: 14px;
            border: 1px solid #bfdbfe;
            background: #eff6ff;
            color: #1d4ed8;
            font-weight: 800;
            min-height: 42px;
        }

        .stTextInput input,
        .stTextArea textarea,
        input,
        textarea {
            background-color: #ffffff !important;
            color: #0f172a !important;
            caret-color: #2563eb !important;
            border-radius: 14px !important;
            border: 1px solid #dbe3ef !important;
            min-height: 42px;
            box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04) !important;
        }

        .stTextInput input:focus,
        .stTextArea textarea:focus,
        input:focus,
        textarea:focus {
            background-color: #ffffff !important;
            color: #0f172a !important;
            border: 1px solid #2563eb !important;
            box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.14) !important;
        }

        .stTextInput input::placeholder,
        .stTextArea textarea::placeholder,
        input::placeholder,
        textarea::placeholder {
            color: #94a3b8 !important;
            opacity: 1 !important;
        }

        .stSelectbox div[data-baseweb="select"] {
            border-radius: 14px !important;
            background-color: #ffffff !important;
            color: #0f172a !important;
        }

        .stSelectbox div[data-baseweb="select"] > div {
            background-color: #ffffff !important;
            color: #0f172a !important;
            border-color: #dbe3ef !important;
            border-radius: 14px !important;
        }

        .stSelectbox div[data-baseweb="select"] span {
            color: #0f172a !important;
        }

        .stSelectbox div[data-baseweb="select"] svg {
            fill: #334155 !important;
            color: #334155 !important;
        }

        div[data-baseweb="popover"] {
            background-color: #ffffff !important;
            color: #0f172a !important;
        }

        div[data-baseweb="popover"] * {
            background-color: #ffffff !important;
            color: #0f172a !important;
        }

        [role="listbox"] {
            background-color: #ffffff !important;
            color: #0f172a !important;
        }

        [role="option"] {
            background-color: #ffffff !important;
            color: #0f172a !important;
        }

        [role="option"]:hover {
            background-color: #eff6ff !important;
            color: #1d4ed8 !important;
        }

        [data-testid="stFileUploaderDropzone"] {
            background-color: #ffffff !important;
            border: 1px solid #e2e8f0 !important;
            border-radius: 16px !important;
        }

        [data-testid="stFileUploaderDropzone"] * {
            color: #334155 !important;
        }

        [data-testid="stTextInputRootElement"],
        [data-testid="stTextAreaRootElement"] {
            background-color: transparent !important;
        }

        [data-testid="stMarkdownContainer"] p {
            color: inherit;
        }

        div[data-testid="stLinkButton"] a,
        .stLinkButton a {
            border-radius: 14px !important;
            background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%) !important;
            color: #ffffff !important;
            border: none !important;
            font-weight: 800 !important;
            min-height: 42px !important;
            box-shadow: 0 10px 22px rgba(37, 99, 235, 0.20) !important;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
        }

        .stTabs [data-baseweb="tab"] {
            border-radius: 999px;
            padding: 8px 18px;
            background: #f1f5f9;
        }

        .stTabs [aria-selected="true"] {
            background: #dbeafe;
            color: #1d4ed8;
            font-weight: 800;
        }

        div[data-testid="stSidebar"] {
            background: rgba(255, 255, 255, 0.86);
            border-right: 1px solid #e5e7eb;
        }

        div[data-testid="stDataFrame"] {
            border-radius: 18px;
            overflow: hidden;
            border: 1px solid #e5e7eb;
        }

        header {
            background: transparent !important;
        }

        footer {
            visibility: hidden;
        }

        #MainMenu {
            visibility: hidden;
        }

        @media (max-width: 900px) {
            .block-container {
                padding-left: 1.2rem;
                padding-right: 1.2rem;
            }

            .landing-title {
                font-size: 42px;
            }

            .landing-hero {
                padding: 36px 30px;
            }
        }
        </style>
        """,
        unsafe_allow_html=True
    )


def render_metric_card(label, value, note=""):
    note_html = f'<div class="metric-note">{note}</div>' if note else ""
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            {note_html}
        </div>
        """,
        unsafe_allow_html=True
    )


def render_section_title(title, subtitle=""):
    subtitle_html = f'<div class="section-subtitle">{subtitle}</div>' if subtitle else ""
    st.markdown(
        f"""
        <div class="section-title">{title}</div>
        {subtitle_html}
        """,
        unsafe_allow_html=True
    )


def render_info_card(title, content):
    st.markdown(
        f"""
        <div class="info-card">
            <div class="info-card-title">{title}</div>
            <div class="info-card-text">{content}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def render_status_message(recommendation):
    if recommendation == "建议投递":
        css_class = "status-success"
        text = "建议投递：这个岗位和求职者背景比较匹配。"
    elif recommendation == "可以投递":
        css_class = "status-warning"
        text = "可以投递：但建议根据岗位重点调整简历和话术。"
    else:
        css_class = "status-danger"
        text = "不太建议投递：匹配度偏低，可能浪费时间。"

    st.markdown(
        f"""
        <div class="status-pill {css_class}">{text}</div>
        """,
        unsafe_allow_html=True
    )


@st.dialog("升级套餐", width="large")
def show_pricing_dialog():
    st.markdown(
        """
        <div class="pricing-header">
            <div class="pricing-title">开通会员套餐</div>
            <div class="pricing-subtitle">
                当前仍处于公测阶段，普通用户暂时不限次数。这里先展示后续正式付费版本的套餐结构。
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    plan_col1, plan_col2, plan_col3 = st.columns(3, gap="large")

    with plan_col1:
        st.markdown(
            f"""
            <div class="plan-card">
                <div class="plan-name">周卡</div>
                <div class="plan-price">{PLAN_PRICES["周卡"]}</div>
                <div class="plan-duration">{PLAN_DURATIONS["周卡"]}</div>
                <div class="plan-feature">✓ 不限岗位分析次数</div>
                <div class="plan-feature">✓ 多份简历智能匹配</div>
                <div class="plan-feature">✓ 求职沟通话术生成</div>
                <div class="plan-feature">✓ 适合短期集中投递</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button("选择周卡", use_container_width=True, key="choose_weekly"):
            st.info("当前公测阶段暂不需要支付。正式上线后，可通过微信/支付宝联系管理员开通周卡。")

    with plan_col2:
        st.markdown(
            f"""
            <div class="plan-card plan-card-featured">
                <div class="plan-badge">推荐</div>
                <div class="plan-name">月卡</div>
                <div class="plan-price">{PLAN_PRICES["月卡"]}</div>
                <div class="plan-duration">{PLAN_DURATIONS["月卡"]}</div>
                <div class="plan-feature">✓ 不限岗位分析次数</div>
                <div class="plan-feature">✓ 多方向简历匹配</div>
                <div class="plan-feature">✓ 投递记录长期管理</div>
                <div class="plan-feature">✓ 适合一个求职周期</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button("选择月卡", use_container_width=True, key="choose_monthly"):
            st.info("当前公测阶段暂不需要支付。正式上线后，可通过微信/支付宝联系管理员开通月卡。")

    with plan_col3:
        st.markdown(
            f"""
            <div class="plan-card">
                <div class="plan-badge">最划算</div>
                <div class="plan-name">年卡</div>
                <div class="plan-price">{PLAN_PRICES["年卡"]}</div>
                <div class="plan-duration">{PLAN_DURATIONS["年卡"]}</div>
                <div class="plan-feature">✓ 全年不限岗位分析</div>
                <div class="plan-feature">✓ 长期保存投递记录</div>
                <div class="plan-feature">✓ 适合长期求职和跳槽准备</div>
                <div class="plan-feature">✓ 单月成本更低</div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button("选择年卡", use_container_width=True, key="choose_yearly"):
            st.info("当前公测阶段暂不需要支付。正式上线后，可通过微信/支付宝联系管理员开通年卡。")

    st.markdown(
        """
        <div class="pay-box">
            <strong>开通方式</strong><br>
            当前为公测阶段，暂不需要实际充值，普通用户也暂时不限次数。<br>
            正式上线后，用户可通过微信 / 支付宝联系管理员付款，并提供自己的用户名；
            管理员在后台将账号设置为周卡、月卡或年卡后，即可获得不限次数权限。
        </div>
        """,
        unsafe_allow_html=True
    )


def render_hero():
    membership = get_user_membership(st.session_state.username)
    membership_type = membership.get("membership_type", "免费试用")

    if PUBLIC_BETA_UNLIMITED:
        status_text = "公测不限次数"
    elif int(membership.get("is_unlimited", 0)) == 1:
        status_text = f"{membership_type} · 不限次数"
    else:
        today_count = get_today_usage_count(st.session_state.username)
        daily_limit = int(membership.get("daily_limit", FREE_DAILY_ANALYSIS_LIMIT))
        status_text = f"免费试用 · 今日 {today_count}/{daily_limit} 次"

    hero_left, hero_right = st.columns([4.3, 1.15], gap="large")

    with hero_left:
        st.markdown(
            """
            <div class="hero-card">
                <div class="hero-title">投递侠AI</div>
                <div class="hero-subtitle">
                    AI 简历匹配与半自动投递工作台。上传多份简历，粘贴岗位 JD，系统会自动分析岗位匹配度、推荐合适简历，并生成求职沟通话术。
                </div>
                <span class="tag">简历匹配</span>
                <span class="tag">岗位分析</span>
                <span class="tag">话术生成</span>
                <span class="tag">投递记录</span>
            </div>
            """,
            unsafe_allow_html=True
        )

    with hero_right:
        st.markdown(
            f"""
            <div class="member-card">
                <div class="member-status-pill">{status_text}</div>
                <div class="member-card-title">升级套餐</div>
                <div class="member-card-text">
                    开通周卡、月卡或年卡后，可获得不限次数的岗位分析与简历匹配权限。
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if st.button("开通周卡 / 月卡 / 年卡", use_container_width=True):
            show_pricing_dialog()


def init_user_db():
    conn = sqlite3.connect(USER_DB_FILE)
    cursor = conn.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS user_memberships (
            username TEXT PRIMARY KEY,
            membership_type TEXT NOT NULL DEFAULT '免费试用',
            daily_limit INTEGER NOT NULL DEFAULT 3,
            is_unlimited INTEGER NOT NULL DEFAULT 0,
            start_date TEXT,
            end_date TEXT,
            updated_at TEXT NOT NULL
        )
        """
    )

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS usage_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL,
            usage_date TEXT NOT NULL,
            action_type TEXT NOT NULL DEFAULT 'analyze_job',
            count INTEGER NOT NULL DEFAULT 0,
            updated_at TEXT NOT NULL,
            UNIQUE(username, usage_date, action_type)
        )
        """
    )

    conn.commit()
    conn.close()


def hash_password(password):
    salt = secrets.token_hex(16)
    password_hash = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        120000
    ).hex()

    return f"{salt}${password_hash}"


def verify_password(password, stored_password_hash):
    try:
        salt, saved_hash = stored_password_hash.split("$")
    except ValueError:
        return False

    new_hash = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        120000
    ).hex()

    return hmac.compare_digest(new_hash, saved_hash)


def username_exists(username):
    conn = sqlite3.connect(USER_DB_FILE)
    cursor = conn.cursor()

    cursor.execute("SELECT id FROM users WHERE username = ?", (username,))
    result = cursor.fetchone()

    conn.close()

    return result is not None


def ensure_user_membership(username):
    username = username.strip()

    if not username:
        return

    conn = sqlite3.connect(USER_DB_FILE)
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT username
        FROM user_memberships
        WHERE username = ?
        """,
        (username,)
    )

    exists = cursor.fetchone()

    if exists is None:
        cursor.execute(
            """
            INSERT INTO user_memberships (
                username,
                membership_type,
                daily_limit,
                is_unlimited,
                start_date,
                end_date,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                username,
                "免费试用",
                FREE_DAILY_ANALYSIS_LIMIT,
                0,
                "",
                "",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            )
        )

    conn.commit()
    conn.close()


def create_user(username, password, admin_register_code=""):
    username = username.strip()

    if not username:
        return False, "用户名不能为空。"

    if len(username) < 3:
        return False, "用户名至少需要 3 个字符。"

    if len(password) < 6:
        return False, "密码至少需要 6 位。"

    if username == ADMIN_USERNAME:
        if not hmac.compare_digest(admin_register_code.strip(), ADMIN_REGISTER_CODE):
            return False, "管理员账号不能通过普通注册入口注册。"

    if username_exists(username):
        return False, "这个用户名已经被注册，请换一个。"

    password_hash = hash_password(password)

    conn = sqlite3.connect(USER_DB_FILE)
    cursor = conn.cursor()

    cursor.execute(
        """
        INSERT INTO users (username, password_hash, created_at)
        VALUES (?, ?, ?)
        """,
        (
            username,
            password_hash,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        )
    )

    conn.commit()
    conn.close()

    ensure_user_membership(username)

    return True, "注册成功，请返回登录。"


def login_user(username, password):
    username = username.strip()

    conn = sqlite3.connect(USER_DB_FILE)
    cursor = conn.cursor()

    cursor.execute(
        "SELECT password_hash FROM users WHERE username = ?",
        (username,)
    )

    result = cursor.fetchone()
    conn.close()

    if result is None:
        return False, "用户名或密码错误。"

    stored_password_hash = result[0]

    if not verify_password(password, stored_password_hash):
        return False, "用户名或密码错误。"

    ensure_user_membership(username)

    return True, "登录成功。"


def is_admin_user(username):
    return hmac.compare_digest(username.strip(), ADMIN_USERNAME)


def get_user_membership(username):
    ensure_user_membership(username)

    conn = sqlite3.connect(USER_DB_FILE)

    membership_df = pd.read_sql_query(
        """
        SELECT username, membership_type, daily_limit, is_unlimited, start_date, end_date, updated_at
        FROM user_memberships
        WHERE username = ?
        """,
        conn,
        params=(username,)
    )

    conn.close()

    if membership_df.empty:
        return {
            "username": username,
            "membership_type": "免费试用",
            "daily_limit": FREE_DAILY_ANALYSIS_LIMIT,
            "is_unlimited": 0,
            "start_date": "",
            "end_date": "",
            "updated_at": "",
        }

    return membership_df.iloc[0].to_dict()


def update_user_membership(username, membership_type, daily_limit=None):
    username = username.strip()
    membership_type = membership_type.strip()

    if membership_type not in MEMBERSHIP_TYPES:
        membership_type = "免费试用"

    is_unlimited = 1 if membership_type in UNLIMITED_MEMBERSHIPS else 0

    if daily_limit is None:
        daily_limit = FREE_DAILY_ANALYSIS_LIMIT

    if is_unlimited:
        daily_limit = 999999

    ensure_user_membership(username)

    conn = sqlite3.connect(USER_DB_FILE)
    cursor = conn.cursor()

    cursor.execute(
        """
        UPDATE user_memberships
        SET membership_type = ?,
            daily_limit = ?,
            is_unlimited = ?,
            updated_at = ?
        WHERE username = ?
        """,
        (
            membership_type,
            int(daily_limit),
            int(is_unlimited),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            username
        )
    )

    conn.commit()
    conn.close()


def get_today_usage_count(username):
    today = datetime.now().strftime("%Y-%m-%d")

    conn = sqlite3.connect(USER_DB_FILE)
    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT count
        FROM usage_records
        WHERE username = ?
          AND usage_date = ?
          AND action_type = 'analyze_job'
        """,
        (username, today)
    )

    result = cursor.fetchone()
    conn.close()

    if result is None:
        return 0

    return int(result[0])


def record_analysis_usage(username):
    today = datetime.now().strftime("%Y-%m-%d")
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = sqlite3.connect(USER_DB_FILE)
    cursor = conn.cursor()

    cursor.execute(
        """
        INSERT INTO usage_records (
            username,
            usage_date,
            action_type,
            count,
            updated_at
        )
        VALUES (?, ?, 'analyze_job', 1, ?)
        ON CONFLICT(username, usage_date, action_type)
        DO UPDATE SET
            count = count + 1,
            updated_at = excluded.updated_at
        """,
        (username, today, now)
    )

    conn.commit()
    conn.close()


def can_user_analyze(username):
    membership = get_user_membership(username)
    today_count = get_today_usage_count(username)

    if PUBLIC_BETA_UNLIMITED:
        return True, "公测阶段不限次数。"

    if int(membership.get("is_unlimited", 0)) == 1:
        return True, "会员不限次数。"

    daily_limit = int(membership.get("daily_limit", FREE_DAILY_ANALYSIS_LIMIT))

    if today_count < daily_limit:
        return True, f"今日剩余 {daily_limit - today_count} 次。"

    return False, f"今日免费分析次数已用完，每天免费 {daily_limit} 次。"


def create_empty_dataframe():
    return pd.DataFrame(columns=ALL_COLUMNS)


def ensure_columns(df):
    for column in ALL_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    return df[ALL_COLUMNS]


def load_all_applications():
    if os.path.exists(DATA_FILE):
        df = pd.read_csv(DATA_FILE, encoding="utf-8-sig")
        return ensure_columns(df)

    return create_empty_dataframe()


def load_user_applications(username):
    df = load_all_applications()

    if "用户名" not in df.columns:
        df["用户名"] = ""

    return df[df["用户名"] == username].copy()


def save_dataframe(df):
    df = ensure_columns(df)
    df.to_csv(DATA_FILE, index=False, encoding="utf-8-sig")


def save_application(record):
    df = load_all_applications()
    new_row = pd.DataFrame([record])
    df = pd.concat([df, new_row], ignore_index=True)
    df = ensure_columns(df)
    save_dataframe(df)


def clear_user_applications(username):
    df = load_all_applications()
    df = df[df["用户名"] != username].copy()
    save_dataframe(df)


def load_users_for_admin():
    if not os.path.exists(USER_DB_FILE):
        return pd.DataFrame(columns=["id", "username", "created_at"])

    conn = sqlite3.connect(USER_DB_FILE)

    users_df = pd.read_sql_query(
        """
        SELECT id, username, created_at
        FROM users
        ORDER BY id DESC
        """,
        conn
    )

    conn.close()

    for username in users_df["username"].tolist():
        ensure_user_membership(username)

    return users_df


def load_memberships_for_admin():
    load_users_for_admin()

    conn = sqlite3.connect(USER_DB_FILE)

    memberships_df = pd.read_sql_query(
        """
        SELECT username, membership_type, daily_limit, is_unlimited, start_date, end_date, updated_at
        FROM user_memberships
        ORDER BY updated_at DESC
        """,
        conn
    )

    conn.close()

    return memberships_df


def load_usage_for_admin():
    conn = sqlite3.connect(USER_DB_FILE)

    usage_df = pd.read_sql_query(
        """
        SELECT username, usage_date, action_type, count, updated_at
        FROM usage_records
        ORDER BY usage_date DESC, count DESC
        """,
        conn
    )

    conn.close()

    return usage_df


def render_admin_dashboard():
    if not is_admin_user(st.session_state.username):
        st.error("你没有权限访问后台管理。")
        st.stop()

    st.markdown(
        """
        <div class="hero-card">
            <div class="hero-title">后台管理</div>
            <div class="hero-subtitle">
                查看注册用户、投递记录、用户权限和每日分析次数。这里只显示用户信息和业务数据，不显示密码。
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    users_df = load_users_for_admin()
    all_applications_df = load_all_applications()
    usage_df = load_usage_for_admin()

    total_users = len(users_df)
    total_applications = len(all_applications_df)

    today = datetime.now().strftime("%Y-%m-%d")

    if not users_df.empty:
        today_users = users_df["created_at"].astype(str).str.startswith(today).sum()
    else:
        today_users = 0

    if not all_applications_df.empty:
        today_applications = all_applications_df["保存时间"].astype(str).str.startswith(today).sum()
    else:
        today_applications = 0

    if not usage_df.empty:
        today_usage_total = usage_df[usage_df["usage_date"].astype(str) == today]["count"].sum()
    else:
        today_usage_total = 0

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        render_metric_card("注册人数", total_users)

    with col2:
        render_metric_card("今日新增用户", today_users)

    with col3:
        render_metric_card("投递记录数", total_applications)

    with col4:
        render_metric_card("今日新增记录", today_applications)

    with col5:
        render_metric_card("今日分析次数", int(today_usage_total))

    st.info(
        "当前处于公测阶段：普通用户暂时不限分析次数，但系统会记录每个用户每天点击“分析这个岗位”的次数。"
    )

    render_section_title("注册用户列表", "这里只显示用户名和注册时间，不显示密码。")

    if users_df.empty:
        st.info("目前还没有注册用户。")
    else:
        st.dataframe(users_df, use_container_width=True)

        users_csv = users_df.to_csv(index=False, encoding="utf-8-sig")

        st.download_button(
            label="下载用户数据 CSV",
            data=users_csv,
            file_name="users_admin_export.csv",
            mime="text/csv"
        )

    render_section_title("用户权限管理", "可以设置用户为免费试用、周卡、月卡、年卡。周卡、月卡、年卡会被系统识别为不限次数。")

    memberships_df = load_memberships_for_admin()

    if memberships_df.empty:
        st.info("目前还没有用户权限数据。")
    else:
        editable_memberships_df = memberships_df.copy()

        editable_memberships_df["是否不限次数"] = editable_memberships_df["is_unlimited"].apply(
            lambda x: "是" if int(x) == 1 else "否"
        )

        display_memberships_df = editable_memberships_df[
            [
                "username",
                "membership_type",
                "daily_limit",
                "是否不限次数",
                "updated_at"
            ]
        ].copy()

        edited_memberships_df = st.data_editor(
            display_memberships_df,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "username": st.column_config.TextColumn("用户名"),
                "membership_type": st.column_config.SelectboxColumn(
                    "会员类型",
                    options=MEMBERSHIP_TYPES,
                    required=True,
                ),
                "daily_limit": st.column_config.NumberColumn(
                    "每日免费次数",
                    min_value=0,
                    step=1,
                ),
                "是否不限次数": st.column_config.TextColumn("是否不限次数"),
                "updated_at": st.column_config.TextColumn("更新时间"),
            },
            disabled=["username", "是否不限次数", "updated_at"],
        )

        if st.button("保存用户权限设置"):
            for _, row in edited_memberships_df.iterrows():
                update_user_membership(
                    username=row["username"],
                    membership_type=row["membership_type"],
                    daily_limit=row["daily_limit"],
                )

            st.success("用户权限设置已保存。")
            st.rerun()

    render_section_title("用户使用次数记录", "这里记录每个用户每天点击“分析这个岗位”的次数。")

    if usage_df.empty:
        st.info("目前还没有使用次数记录。")
    else:
        st.dataframe(usage_df, use_container_width=True)

        usage_csv = usage_df.to_csv(index=False, encoding="utf-8-sig")

        st.download_button(
            label="下载使用次数记录 CSV",
            data=usage_csv,
            file_name="usage_records_admin_export.csv",
            mime="text/csv"
        )

    render_section_title("全部投递记录", "这里显示所有账号保存过的投递记录。")

    if all_applications_df.empty:
        st.info("目前还没有任何投递记录。")
    else:
        st.dataframe(all_applications_df, use_container_width=True)

        applications_csv = all_applications_df.to_csv(index=False, encoding="utf-8-sig")

        st.download_button(
            label="下载全部投递记录 CSV",
            data=applications_csv,
            file_name="all_applications_admin_export.csv",
            mime="text/csv"
        )


def show_auth_page():
    st.markdown(
        """
        <div class="auth-brand-row">
            <div class="auth-logo">
                <span class="auth-logo-icon">侠</span>
                <span>投递侠AI</span>
            </div>
            <div class="auth-version">AI 简历匹配与半自动投递工作台</div>
        </div>
        """,
        unsafe_allow_html=True
    )

    left_col, right_col = st.columns([1.45, 0.95], gap="large")

    with left_col:
        st.markdown(
            """
            <div class="landing-hero">
                <div class="landing-kicker">为正在求职的人，减少无效投递</div>
                <div class="landing-title">让每一次简历投递，都更精准一点。</div>
                <div class="landing-subtitle">
                    上传多份简历，粘贴岗位 JD，投递侠AI 会自动识别岗位方向、分析匹配度、
                    推荐更适合投递的简历版本，并生成自然的求职沟通话术。
                    你可以把它当成自己的求职投递工作台。
                </div>
                <div class="landing-tags">
                    <span class="landing-tag">多份简历智能匹配</span>
                    <span class="landing-tag">岗位 JD 自动分析</span>
                    <span class="landing-tag">求职话术生成</span>
                    <span class="landing-tag">投递记录管理</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with right_col:
        st.markdown(
            """
            <div class="login-card-title">账号入口</div>
            <div class="login-card-subtitle">
                注册后即可开始使用。上传简历、分析岗位、管理投递记录，都可以在一个页面完成。
            </div>
            """,
            unsafe_allow_html=True
        )

        tab_login, tab_register = st.tabs(["登录", "注册"])

        with tab_login:
            with st.form("login_form"):
                username = st.text_input("用户名")
                password = st.text_input("密码", type="password")
                submitted = st.form_submit_button("登录")

                if submitted:
                    success, message = login_user(username, password)

                    if success:
                        st.session_state.logged_in = True
                        st.session_state.username = username.strip()
                        st.session_state.latest_result = None
                        st.success(message)
                        st.rerun()
                    else:
                        st.error(message)

        with tab_register:
            with st.form("register_form"):
                new_username = st.text_input("设置用户名")
                new_password = st.text_input("设置密码", type="password")
                confirm_password = st.text_input("确认密码", type="password")

                admin_register_code = st.text_input(
                    "管理员注册密钥",
                    type="password",
                    help="普通用户不用填写。只有注册管理员账号时需要。"
                )

                submitted = st.form_submit_button("注册")

                if submitted:
                    if new_password != confirm_password:
                        st.error("两次输入的密码不一致。")
                    else:
                        success, message = create_user(
                            new_username,
                            new_password,
                            admin_register_code
                        )

                        if success:
                            st.success(message)
                        else:
                            st.error(message)

    st.markdown('<div class="feature-grid-title">投递侠AI 适合这些人</div>', unsafe_allow_html=True)

    people_col1, people_col2, people_col3 = st.columns(3, gap="large")

    with people_col1:
        st.markdown(
            """
            <div class="value-card">
                <div class="value-title">应届生 / 留学生</div>
                <div class="value-text">
                    面对多个岗位方向，不知道该用哪份简历投递。系统可以根据岗位 JD，判断用户上传的不同版本简历中，哪一份更适合当前岗位。
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with people_col2:
        st.markdown(
            """
            <div class="value-card">
                <div class="value-title">正在海投的人</div>
                <div class="value-text">
                    每天投递很多岗位，容易忘记投过哪个公司、岗位状态如何、HR 有没有回复。投递侠AI 可以帮你统一记录和跟进。
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with people_col3:
        st.markdown(
            """
            <div class="value-card">
                <div class="value-title">多方向求职者</div>
                <div class="value-text">
                    同时投递多个不同方向的岗位时，可以快速判断岗位重点，减少乱投、错投和低质量投递。
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    scenario_html = (
        '<div class="scenario-list">'
        '<div class="scenario-title">它可以帮你完成什么？</div>'
        '<div class="scenario-item">'
        '<div class="scenario-dot"></div>'
        '<div>把岗位 JD 自动拆解成岗位方向、核心关键词、潜在风险和简历优化建议。</div>'
        '</div>'
        '<div class="scenario-item">'
        '<div class="scenario-dot"></div>'
        '<div>在多份简历中推荐最适合当前岗位的一份，避免每次都凭感觉选择简历。</div>'
        '</div>'
        '<div class="scenario-item">'
        '<div class="scenario-dot"></div>'
        '<div>生成更像真人求职者发给 HR 的沟通话术，而不是机械模板。</div>'
        '</div>'
        '<div class="scenario-item">'
        '<div class="scenario-dot"></div>'
        '<div>保存每一次投递记录，包括公司、岗位、薪资、链接、状态、备注和后续跟进。</div>'
        '</div>'
        '</div>'
    )

    st.markdown(scenario_html, unsafe_allow_html=True)


def extract_text_from_uploaded_file(uploaded_file):
    file_name = uploaded_file.name
    file_name_lower = file_name.lower()
    file_bytes = uploaded_file.getvalue()

    try:
        if file_name_lower.endswith(".txt"):
            try:
                return file_bytes.decode("utf-8"), ""
            except UnicodeDecodeError:
                return file_bytes.decode("gbk", errors="ignore"), ""

        if file_name_lower.endswith(".pdf"):
            if PdfReader is None:
                return "", "缺少 pypdf，请先运行：python -m pip install pypdf"

            reader = PdfReader(io.BytesIO(file_bytes))
            text_list = []

            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text_list.append(page_text)

            text = "\n".join(text_list).strip()

            if not text:
                return "", "PDF 未读取到文字，可能是扫描版或图片版简历。"

            return text, ""

        if file_name_lower.endswith(".docx"):
            if Document is None:
                return "", "缺少 python-docx，请先运行：python -m pip install python-docx"

            document = Document(io.BytesIO(file_bytes))
            text = "\n".join([p.text for p in document.paragraphs]).strip()

            if not text:
                return "", "Word 文件未读取到文字。"

            return text, ""

        return "", "暂不支持该文件类型。"

    except Exception as error:
        return "", f"解析失败：{error}"


def find_keywords(text, keywords):
    text_lower = text.lower()
    matched = []

    for keyword in keywords:
        keyword = keyword.strip()
        if keyword and keyword.lower() in text_lower:
            matched.append(keyword)

    return list(dict.fromkeys(matched))


def analyze_text_categories(text):
    category_result = {}

    for category, keywords in CATEGORY_KEYWORDS.items():
        matched = find_keywords(text, keywords)
        category_result[category] = {
            "count": len(matched),
            "keywords": matched,
        }

    sorted_result = sorted(
        category_result.items(),
        key=lambda item: item[1]["count"],
        reverse=True
    )

    return sorted_result


def get_top_categories(text, limit=3):
    sorted_result = analyze_text_categories(text)
    top = [
        {
            "category": category,
            "count": detail["count"],
            "keywords": detail["keywords"],
        }
        for category, detail in sorted_result
        if detail["count"] > 0
    ]

    return top[:limit]


def calculate_job_score(jd_text, candidate_identity, target_direction):
    score = 55
    reasons = []

    top_categories = get_top_categories(jd_text, limit=3)

    if top_categories:
        main_category = top_categories[0]["category"]
        score += 12
        reasons.append(f"岗位核心方向较明显，主要偏向：{main_category}。")

    fresh_keywords = find_keywords(jd_text, FRESH_GRAD_KEYWORDS)
    if fresh_keywords:
        score += 12
        reasons.append("岗位对应届生、实习生、管培生或低经验候选人较友好。")

    risk_keywords = find_keywords(jd_text, RISK_KEYWORDS)
    if risk_keywords:
        score -= min(20, len(risk_keywords) * 5)
        reasons.append("岗位存在一些需要谨慎判断的关键词：" + "、".join(risk_keywords))

    if "本科" in jd_text or "硕士" in jd_text or "大专" in jd_text:
        score += 5
        reasons.append("岗位描述中有明确学历要求，需要结合求职者学历进一步判断。")

    if target_direction.strip():
        target_keywords = [
            item.strip()
            for item in target_direction.replace("、", ",").replace("，", ",").split(",")
            if item.strip()
        ]

        direction_matches = find_keywords(jd_text, target_keywords)

        if direction_matches:
            score += 8
            reasons.append("岗位内容与求职者主要求职方向存在重合。")

    if "销售" in jd_text and "销售" not in target_direction and "客户" not in target_direction:
        score -= 10
        reasons.append("岗位可能包含销售属性，如果求职者不考虑销售，需要谨慎。")

    if "应届" in candidate_identity and (
        "3年" in jd_text or "三年" in jd_text or "5年" in jd_text or "五年" in jd_text
    ):
        score -= 10
        reasons.append("求职者为应届生，但岗位可能有较高经验要求。")

    score = max(0, min(score, 100))

    if not reasons:
        reasons.append("暂时没有识别到特别明显的岗位匹配点，需要人工进一步判断。")

    return score, reasons


def get_recommendation(score):
    if score >= 80:
        return "建议投递"
    if score >= 65:
        return "可以投递"
    return "不太建议投递"


def score_resume_against_jd(resume_name, resume_text, jd_text):
    resume_full_text = f"{resume_name}\n{resume_text}"
    resume_lower = resume_full_text.lower()

    jd_top_categories = get_top_categories(jd_text, limit=4)

    score = 35
    category_matches = []
    all_matched_keywords = []

    for item in jd_top_categories:
        category = item["category"]
        jd_keywords = item["keywords"]
        matched = find_keywords(resume_lower, jd_keywords)

        if matched:
            category_score = min(25, len(matched) * 5)
            score += category_score
            category_matches.append(f"{category}：命中 {len(matched)} 个关键词")
            all_matched_keywords.extend(matched)

    filename_boost_keywords = []

    for item in jd_top_categories:
        for keyword in item["keywords"]:
            if keyword.lower() in resume_name.lower():
                filename_boost_keywords.append(keyword)

    if filename_boost_keywords:
        score += min(12, len(filename_boost_keywords) * 4)

    if len(resume_text.strip()) >= 300:
        score += 5

    if any(word in resume_lower for word in ["实习", "项目", "负责", "参与", "协助"]):
        score += 5

    if any(word in resume_lower for word in ["数据", "指标", "增长", "转化", "复盘", "统计"]):
        score += 5

    score = max(0, min(score, 100))

    all_matched_keywords = list(dict.fromkeys(all_matched_keywords))
    filename_boost_keywords = list(dict.fromkeys(filename_boost_keywords))

    if all_matched_keywords:
        reason = "简历内容命中岗位关键词：" + "、".join(all_matched_keywords)
    else:
        reason = "简历内容与岗位关键词重合较少。"

    if filename_boost_keywords:
        reason += "；简历文件名也命中：" + "、".join(filename_boost_keywords)

    if category_matches:
        reason += "；方向匹配：" + "；".join(category_matches)

    return {
        "简历文件": resume_name,
        "匹配分": score,
        "命中关键词": "、".join(all_matched_keywords) if all_matched_keywords else "无",
        "方向匹配": "；".join(category_matches) if category_matches else "不明显",
        "推荐依据": reason,
    }


def recommend_resume_from_uploaded_files(resume_items, jd_text):
    if not resume_items:
        return "未上传简历", []

    results = []

    for resume in resume_items:
        result = score_resume_against_jd(
            resume_name=resume["name"],
            resume_text=resume["text"],
            jd_text=jd_text,
        )

        if resume["error"]:
            result["匹配分"] = max(0, result["匹配分"] - 10)
            result["推荐依据"] += f"；注意：{resume['error']}"

        results.append(result)

    results = sorted(results, key=lambda x: x["匹配分"], reverse=True)
    recommended_resume = results[0]["简历文件"]

    return recommended_resume, results


def recommend_resume_by_name_only(jd_text, resume_versions):
    if not resume_versions:
        return "未填写简历版本"

    jd_top_categories = get_top_categories(jd_text, limit=3)

    for item in jd_top_categories:
        category = item["category"]
        keywords = item["keywords"]

        for resume in resume_versions:
            resume_lower = resume.lower()

            if any(keyword.lower() in resume_lower for keyword in keywords):
                return resume

            if "人力" in category and ("hr" in resume_lower or "人力" in resume or "招聘" in resume):
                return resume

            if "运营" in category and ("运营" in resume or "内容" in resume or "新媒体" in resume):
                return resume

            if "市场" in category and ("市场" in resume or "品牌" in resume or "广告" in resume):
                return resume

            if "行政" in category and ("行政" in resume or "助理" in resume or "通用" in resume):
                return resume

    return resume_versions[-1]


def build_job_profile(jd_text):
    top_categories = get_top_categories(jd_text, limit=3)

    if not top_categories:
        return "该岗位方向不够明确，需要人工结合岗位职责进一步判断。"

    profile_parts = []

    for item in top_categories:
        keywords = "、".join(item["keywords"][:8])
        profile_parts.append(f"{item['category']}方向较明显，相关关键词包括：{keywords}")

    return "；".join(profile_parts) + "。"


def build_strengths(recommended_resume, resume_match_results, job_score):
    strengths = []

    if resume_match_results:
        best = resume_match_results[0]
        strengths.append(f"系统推荐使用「{recommended_resume}」，因为它在所有上传简历中匹配分最高，为 {best['匹配分']} 分。")

        if best["命中关键词"] != "无":
            strengths.append(f"该简历与岗位JD存在明显关键词重合，包括：{best['命中关键词']}。")

        if best["方向匹配"] != "不明显":
            strengths.append(f"从方向上看，该简历的匹配点主要体现在：{best['方向匹配']}。")
    else:
        strengths.append(f"系统推荐使用「{recommended_resume}」，当前依据主要来自简历版本名称与岗位关键词的匹配。")

    if job_score >= 80:
        strengths.append("岗位整体匹配度较高，可以优先投递。")
    elif job_score >= 65:
        strengths.append("岗位具备一定匹配度，可以投递，但建议优化话术和简历重点。")
    else:
        strengths.append("岗位整体匹配度一般，建议谨慎投递或作为补充投递。")

    return strengths


def build_risks(jd_text, candidate_identity, target_direction):
    risks = []

    risk_keywords = find_keywords(jd_text, RISK_KEYWORDS)

    if risk_keywords:
        risks.append("岗位中出现以下风险或压力关键词：" + "、".join(risk_keywords) + "。")

    if "销售" in jd_text and "销售" not in target_direction and "客户" not in target_direction:
        risks.append("岗位可能带有销售、客户开发或业绩导向属性，如果求职者不接受销售性质，需要谨慎。")

    if "应届" in candidate_identity and (
        "3年" in jd_text or "三年" in jd_text or "5年" in jd_text or "五年" in jd_text
    ):
        risks.append("求职者身份偏应届，但岗位可能要求多年经验，面试通过率可能受影响。")

    if "独立负责" in jd_text or "独立完成" in jd_text:
        risks.append("岗位可能希望候选人有较强独立负责能力，低经验候选人需要用项目经历补足。")

    if not risks:
        risks.append("未识别到明显高风险点，但仍建议人工确认薪资、工作强度、岗位是否偏销售以及是否接受应届生。")

    return risks


def build_resume_suggestions(jd_text, recommended_resume, resume_match_results):
    suggestions = []
    top_categories = get_top_categories(jd_text, limit=3)

    if top_categories:
        main = top_categories[0]
        keywords = main["keywords"][:8]
        suggestions.append(f"建议在「{recommended_resume}」中优先突出 {main['category']} 相关经历。")
        suggestions.append("建议简历中重点呈现这些岗位关键词：" + "、".join(keywords) + "。")

    if resume_match_results:
        best = resume_match_results[0]
        if best["匹配分"] < 70:
            suggestions.append("当前推荐简历的匹配分不算高，建议针对该岗位单独修改简历关键词。")
        else:
            suggestions.append("当前推荐简历匹配度较好，投递前可微调项目描述，让经历更贴近岗位职责。")

    if any(word in jd_text for word in ["数据", "复盘", "指标", "转化率", "阅读量", "互动率"]):
        suggestions.append("岗位重视数据或复盘，建议简历中补充具体数据结果，例如阅读量、转化率、活动参与人数或效率提升。")

    if any(word in jd_text for word in ["文案", "内容", "小红书", "抖音", "公众号", "短视频"]):
        suggestions.append("岗位重视内容能力，建议简历中加入具体内容平台、选题案例、文案作品或账号运营成果。")

    if any(word in jd_text for word in ["招聘", "面试", "候选人", "人事"]):
        suggestions.append("岗位重视招聘或人事能力，建议简历中突出简历筛选、电话沟通、面试邀约、数据表格等经历。")

    if len(suggestions) == 0:
        suggestions.append("建议根据岗位职责手动调整简历中的项目顺序和关键词。")

    return suggestions


def build_greeting(candidate_identity, education_background, experience_summary, job_title, recommended_resume):
    education_text = education_background.strip()
    experience_text = experience_summary.strip()

    if "市场营销" in education_text:
        major_text = "市场营销专业"
    elif "人力" in education_text or "HR" in education_text or "hr" in education_text:
        major_text = "人力资源相关专业"
    elif "行政" in education_text:
        major_text = "行政管理相关专业"
    elif "传媒" in education_text or "新闻" in education_text or "传播" in education_text:
        major_text = "传媒传播相关专业"
    elif "计算机" in education_text or "软件" in education_text or "数据" in education_text:
        major_text = "计算机或数据相关专业"
    else:
        major_text = "相关专业背景"

    if (
        "新媒体" in experience_text
        or "内容" in experience_text
        or "小红书" in experience_text
        or "抖音" in experience_text
        or "公众号" in experience_text
    ):
        experience_short = "新媒体运营和内容策划相关经历"
    elif "招聘" in experience_text or "人力" in experience_text or "HR" in experience_text or "hr" in experience_text:
        experience_short = "招聘和人力资源相关经历"
    elif "行政" in experience_text or "助理" in experience_text or "文员" in experience_text:
        experience_short = "行政助理和事务支持相关经历"
    elif "市场" in experience_text or "品牌" in experience_text or "广告" in experience_text or "活动" in experience_text:
        experience_short = "市场、品牌和活动执行相关经历"
    elif "数据" in experience_text or "分析" in experience_text or "Excel" in experience_text:
        experience_short = "数据整理和分析相关经历"
    elif "项目" in experience_text or "产品" in experience_text:
        experience_short = "项目协作和产品相关经历"
    else:
        experience_short = "相关实习和项目经历"

    return f"您好，我是{candidate_identity}，{major_text}，有{experience_short}。对贵公司的{job_title}岗位很感兴趣，方便的话能否把简历发您看看？"


def list_to_text(items):
    return "\n".join([f"{index + 1}. {item}" for index, item in enumerate(items)])


def get_count_by_status(df, status):
    if df.empty or "投递状态" not in df.columns:
        return 0
    return int((df["投递状态"] == status).sum())


inject_custom_css()
init_user_db()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "username" not in st.session_state:
    st.session_state.username = ""

if "latest_result" not in st.session_state:
    st.session_state.latest_result = None

if not st.session_state.logged_in:
    show_auth_page()
    st.stop()


st.sidebar.markdown(
    f"""
    <div class="user-badge">
        当前登录用户：{st.session_state.username}
    </div>
    """,
    unsafe_allow_html=True
)

if st.sidebar.button("退出登录"):
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.latest_result = None
    st.rerun()

if is_admin_user(st.session_state.username):
    st.sidebar.divider()
    st.sidebar.success("管理员后台")
    render_admin_dashboard()
    st.stop()

render_hero()

st.sidebar.header("求职者信息")

candidate_identity = st.sidebar.text_input(
    "求职者身份",
    placeholder="例如：本科应届生 / 硕士应届生 / 1年工作经验"
)

education_background = st.sidebar.text_input(
    "教育背景",
    placeholder="例如：本科市场营销，硕士传媒管理"
)

experience_summary = st.sidebar.text_area(
    "经历概述",
    placeholder="例如：有运营实习、社群运营、内容策划和数据整理经验",
    height=120
)

target_direction = st.sidebar.text_input(
    "主要求职方向",
    placeholder="例如：运营、HR、产品、行政"
)

st.sidebar.subheader("简历文件")

uploaded_resumes = st.sidebar.file_uploader(
    "上传简历文件，可多选",
    type=["txt", "pdf", "docx"],
    accept_multiple_files=True
)

resume_versions_text = st.sidebar.text_area(
    "可用简历版本，每行一个；如果已经上传简历，可不填",
    placeholder="例如：\nHR / 人力资源版简历\n运营 / 客户执行版简历\n通用版简历",
    height=100
)

resume_versions = [
    item.strip()
    for item in resume_versions_text.split("\n")
    if item.strip()
]

resume_items = []

if uploaded_resumes:
    st.sidebar.write("已上传简历：")

    for uploaded_file in uploaded_resumes:
        text, error = extract_text_from_uploaded_file(uploaded_file)

        resume_items.append(
            {
                "name": uploaded_file.name,
                "text": text,
                "error": error,
            }
        )

        if error:
            st.sidebar.warning(f"{uploaded_file.name}：{error}")
        else:
            st.sidebar.success(f"{uploaded_file.name}：已读取")

render_section_title("岗位信息", "粘贴岗位 JD 后，系统会自动分析岗位方向、匹配度和推荐简历。")

job_col1, job_col2 = st.columns(2)

with job_col1:
    job_title = st.text_input("岗位名称")
    city = st.text_input("城市")
    job_link = st.text_input("岗位链接")

with job_col2:
    company_name = st.text_input("公司名称")
    salary = st.text_input("薪资")
    application_status = st.selectbox(
        "投递状态",
        ["未投递", "已投递", "已回复", "已面试", "已拒绝", "已录用", "暂不投递"]
    )

note = st.text_area(
    "备注",
    placeholder="例如：岗位偏内容运营；HR回复较快；薪资偏低；需要后续跟进",
    height=80
)

if job_link.strip():
    st.link_button("打开岗位链接", job_link)

jd_text = st.text_area("粘贴岗位JD", height=300)

if st.button("分析这个岗位"):
    has_resume_source = bool(resume_items) or bool(resume_versions_text.strip())

    required_fields = [
        candidate_identity,
        education_background,
        experience_summary,
        target_direction,
        job_title,
        company_name,
        jd_text,
    ]

    if not all(field.strip() for field in required_fields):
        st.warning("请先补全左侧求职者信息，以及右侧岗位名称、公司名称和岗位JD。")
    elif not has_resume_source:
        st.warning("请至少上传一份简历，或手动填写可用简历版本。")
    else:
        can_analyze, limit_message = can_user_analyze(st.session_state.username)

        if not can_analyze:
            st.error(limit_message)
            st.stop()

        record_analysis_usage(st.session_state.username)

        job_score, reasons = calculate_job_score(
            jd_text=jd_text,
            candidate_identity=candidate_identity,
            target_direction=target_direction,
        )

        recommendation = get_recommendation(job_score)

        if resume_items:
            recommended_resume, resume_match_results = recommend_resume_from_uploaded_files(
                resume_items=resume_items,
                jd_text=jd_text,
            )
        else:
            recommended_resume = recommend_resume_by_name_only(jd_text, resume_versions)
            resume_match_results = []

        job_profile = build_job_profile(jd_text)
        strengths = build_strengths(recommended_resume, resume_match_results, job_score)
        risks = build_risks(jd_text, candidate_identity, target_direction)
        suggestions = build_resume_suggestions(jd_text, recommended_resume, resume_match_results)

        greeting = build_greeting(
            candidate_identity=candidate_identity,
            education_background=education_background,
            experience_summary=experience_summary,
            job_title=job_title,
            recommended_resume=recommended_resume,
        )

        st.session_state.latest_result = {
            "job_score": job_score,
            "reasons": reasons,
            "recommendation": recommendation,
            "recommended_resume": recommended_resume,
            "resume_match_results": resume_match_results,
            "job_profile": job_profile,
            "strengths": strengths,
            "risks": risks,
            "suggestions": suggestions,
            "greeting": greeting,
        }

        if PUBLIC_BETA_UNLIMITED:
            st.info("公测阶段当前不限分析次数，本次分析已记录到后台。")


if st.session_state.latest_result:
    result = st.session_state.latest_result

    render_section_title("岗位分析结果", "系统会综合岗位 JD、求职方向和简历内容，给出本次投递建议。")

    col1, col2, col3 = st.columns(3)

    with col1:
        render_metric_card("岗位匹配度", f"{result['job_score']}/100", "分数越高，越建议优先投递")

    with col2:
        render_metric_card("推荐结论", result["recommendation"], "结合岗位要求与候选人信息判断")

    with col3:
        render_metric_card("推荐简历", result["recommended_resume"], "建议本次投递优先使用")

    render_status_message(result["recommendation"])

    left_result, right_result = st.columns([1.05, 0.95], gap="large")

    with left_result:
        render_info_card("岗位画像", result["job_profile"])

        render_section_title("简历推荐结果")

        st.write(f"建议使用：**{result['recommended_resume']}**")

        if result["resume_match_results"]:
            st.write("各简历与岗位 JD 的匹配情况：")
            st.dataframe(
                pd.DataFrame(result["resume_match_results"]),
                use_container_width=True
            )
        else:
            st.info("当前未上传简历文件，系统根据手动填写的简历版本名称进行推荐。")

    with right_result:
        render_section_title("推荐打招呼语")
        st.write("下面这段可以直接复制到招聘平台：")
        st.code(result["greeting"], language="text")
        st.text_area("可复制话术", result["greeting"], height=120)

        st.markdown("**投递动作提醒**")
        st.write(f"建议上传或选择这份简历：**{result['recommended_resume']}**")

        if job_link.strip():
            st.link_button("去招聘平台投递这个岗位", job_link)

    analysis_col1, analysis_col2 = st.columns(2, gap="large")

    with analysis_col1:
        render_section_title("匹配优势")
        for item in result["strengths"]:
            st.write(f"- {item}")

        render_section_title("岗位分析理由")
        for reason in result["reasons"]:
            st.write(f"- {reason}")

    with analysis_col2:
        render_section_title("风险点")
        for item in result["risks"]:
            st.write(f"- {item}")

        render_section_title("简历优化建议")
        for item in result["suggestions"]:
            st.write(f"- {item}")

    if st.button("保存到投递记录"):
        if result["resume_match_results"]:
            resume_detail = "；".join(
                [
                    f"{item['简历文件']}：{item['匹配分']}分，命中关键词：{item['命中关键词']}"
                    for item in result["resume_match_results"]
                ]
            )
        else:
            resume_detail = "未上传简历文件，仅根据简历版本名称推荐。"

        record = {
            "用户名": st.session_state.username,
            "保存时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "求职者身份": candidate_identity,
            "教育背景": education_background,
            "经历概述": experience_summary,
            "主要求职方向": target_direction,
            "公司名称": company_name,
            "岗位名称": job_title,
            "城市": city,
            "薪资": salary,
            "岗位链接": job_link,
            "匹配度评分": result["job_score"],
            "推荐结论": result["recommendation"],
            "推荐简历版本": result["recommended_resume"],
            "简历匹配详情": resume_detail,
            "岗位画像": result["job_profile"],
            "匹配优势": list_to_text(result["strengths"]),
            "风险点": list_to_text(result["risks"]),
            "简历优化建议": list_to_text(result["suggestions"]),
            "投递状态": application_status,
            "备注": note,
            "打招呼语": result["greeting"],
            "岗位JD": jd_text,
        }

        save_application(record)
        st.success("已保存到投递记录。")


st.divider()

render_section_title("历史投递记录", "当前账号的投递记录、状态统计和后续跟进都在这里管理。")

applications_df = load_user_applications(st.session_state.username)

if applications_df.empty:
    st.info("目前还没有保存任何投递记录。")
else:
    total_count = len(applications_df)
    applied_count = get_count_by_status(applications_df, "已投递")
    replied_count = get_count_by_status(applications_df, "已回复")
    interview_count = get_count_by_status(applications_df, "已面试")

    m1, m2, m3, m4 = st.columns(4)

    with m1:
        render_metric_card("总记录", total_count)

    with m2:
        render_metric_card("已投递", applied_count)

    with m3:
        render_metric_card("已回复", replied_count)

    with m4:
        render_metric_card("已面试", interview_count)

    status_counts = applications_df["投递状态"].fillna("未填写").value_counts()

    render_section_title("状态统计")
    st.bar_chart(status_counts)

    render_section_title("投递记录表")

    display_columns = [
        "保存时间",
        "公司名称",
        "岗位名称",
        "城市",
        "薪资",
        "匹配度评分",
        "推荐结论",
        "推荐简历版本",
        "投递状态",
        "备注",
        "岗位链接",
    ]

    display_df = applications_df[display_columns].copy()

    edited_df = st.data_editor(
        display_df,
        use_container_width=True,
        num_rows="fixed",
        column_config={
            "投递状态": st.column_config.SelectboxColumn(
                "投递状态",
                options=["未投递", "已投递", "已回复", "已面试", "已拒绝", "已录用", "暂不投递"],
                required=True,
            ),
            "备注": st.column_config.TextColumn(
                "备注",
                help="可以记录HR反馈、岗位风险、后续跟进事项等。",
            ),
            "岗位链接": st.column_config.LinkColumn(
                "岗位链接",
                help="点击打开岗位页面。",
            ),
        },
        disabled=[
            "保存时间",
            "公司名称",
            "岗位名称",
            "城市",
            "薪资",
            "匹配度评分",
            "推荐结论",
            "推荐简历版本",
            "岗位链接",
        ],
    )

    if st.button("保存历史记录修改"):
        all_df = load_all_applications()

        for index in edited_df.index:
            all_df.loc[index, "投递状态"] = edited_df.loc[index, "投递状态"]
            all_df.loc[index, "备注"] = edited_df.loc[index, "备注"]

        save_dataframe(all_df)
        st.success("历史记录修改已保存。")

    csv_data = applications_df.to_csv(index=False, encoding="utf-8-sig")

    st.download_button(
        label="下载我的投递记录 CSV",
        data=csv_data,
        file_name=f"{st.session_state.username}_applications.csv",
        mime="text/csv"
    )

    with st.expander("查看完整字段"):
        st.dataframe(applications_df, use_container_width=True)

    with st.expander("危险操作：清空当前账号的全部投递记录"):
        confirm_clear = st.checkbox("我确认要清空当前账号的全部投递记录")
        if confirm_clear:
            if st.button("清空我的全部记录"):
                clear_user_applications(st.session_state.username)
                st.success("当前账号的全部记录已清空，请刷新页面查看。")