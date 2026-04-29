import streamlit as st
from openai import OpenAI
import re
import pandas as pd
from io import BytesIO
from datetime import datetime, date
import plotly.graph_objects as go
import plotly.express as px
from streamlit import cache_data, cache_resource
import json

# ═══════════════════════════════════════════════════════════════════════
# API 配置
# ═══════════════════════════════════════════════════════════════════════
API_KEY = "sk-e8f4371072b84703b005e2445b1fc24e"
client = OpenAI(api_key=API_KEY, base_url="https://api.deepseek.com")

st.set_page_config(
    page_title="校园财务问答助手",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ═══════════════════════════════════════════════════════════════════════
# 全局样式（浅青色主题 + 消息左右对齐）
# ═══════════════════════════════════════════════════════════════════════
@cache_resource
def get_css():
    return """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@600;700&family=Noto+Sans+SC:wght@300;400;500;700&display=swap');
    :root {
        --primary: #6DC6C9;
        --primary-dark: #1B8B8F;
        --primary-light: #E8F4F5;
        --primary-bright: #1FB5BA;
        --accent: #2B4B6E;
        --accent-light: #4A6F8E;
        --bg-body: #F4FAFB;
        --bg-card: #FFFFFF;
        --text: #1A2C3C;
        --text-muted: #5A6E7A;
        --border: #CDDFE3;
    }
    .stApp { background: var(--bg-body); color: var(--text); font-family: 'Noto Sans SC', sans-serif; }
    header[data-testid="stHeader"] { background: rgba(244,250,251,0.96); backdrop-filter: blur(4px); border-bottom: 1px solid var(--border); }
    section[data-testid="stSidebar"] { background: #FFFFFF; border-right: 1px solid var(--border); }
    .hero {
        background: linear-gradient(125deg, #FFFFFF 0%, var(--primary-light) 100%);
        border: 1px solid var(--border);
        border-radius: 18px;
        padding: 36px 44px;
        margin-bottom: 26px;
        position: relative;
        overflow: hidden;
    }
    .hero-title { font-family: 'Noto Serif SC', serif; font-size: 2.2rem; font-weight: 700; color: var(--accent); letter-spacing: 3px; margin: 0 0 10px; }
    .hero-sub { color: var(--text-muted); font-size: 0.92rem; margin: 0 0 16px; font-weight: 300; }
    .badge {
        background: var(--primary-light);
        border: 1px solid var(--primary);
        color: var(--primary-dark);
        border-radius: 20px;
        padding: 3px 13px;
        font-size: 0.73rem;
        margin: 3px 6px 3px 0;
    }
    .bubble-user {
        background: var(--primary-light);
        border-radius: 14px 14px 4px 14px;
        padding: 12px 18px;
        margin: 6px 0 6px auto;
        max-width: 70%;
        font-size: 0.9rem;
        line-height: 1.8;
        color: var(--text);
        text-align: right;
    }
    .bubble-ai {
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: 14px 14px 14px 4px;
        padding: 12px 18px;
        margin: 6px auto 6px 0;
        max-width: 70%;
        font-size: 0.9rem;
        line-height: 1.8;
        color: var(--text);
        text-align: left;
    }
    .bubble-label {
        font-size: 0.7rem;
        color: var(--text-muted);
        margin-bottom: 5px;
    }
    .user-label { text-align: right; }
    .ai-label { text-align: left; }
    .chat-empty { text-align: center; color: #B0C4CE; padding: 40px; font-size: 0.9rem; }
    .quick-chip {
        background: var(--primary-light);
        border: 1px solid var(--border);
        color: var(--primary-dark);
        border-radius: 22px;
        padding: 5px 15px;
        font-size: 0.78rem;
        cursor: pointer;
        transition: all 0.2s;
    }
    .quick-chip:hover { background: var(--primary); color: white; border-color: var(--primary); }
    .cl {
        background: var(--bg-card);
        border-left: 3px solid var(--primary);
        border-radius: 0 9px 9px 0;
        padding: 10px 16px;
        margin: 5px 0;
        font-size: 0.87rem;
        color: var(--text);
    }
    .cl.warn { border-left-color: #E74C3C; background: #FFF4F2; color: #C0392B; }
    .cl.ok { border-left-color: #27AE60; background: #F0FFF0; color: #1E6F3F; }
    .cl.info { border-left-color: var(--primary); background: var(--primary-light); color: var(--primary-dark); }
    .divider { height: 1px; margin: 20px 0; background: linear-gradient(90deg, transparent, var(--border), transparent); }
    .stTextInput > div > div > input, .stTextArea > div > div > textarea, .stNumberInput > div > div > input {
        background: var(--bg-card) !important;
        border: 1px solid var(--border) !important;
        border-radius: 10px !important;
        color: var(--text) !important;
    }
    .stButton > button {
        background: var(--accent) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        transition: all 0.25s !important;
    }
    .stButton > button:hover { background: var(--accent-light) !important; box-shadow: 0 4px 12px rgba(43,75,110,0.3) !important; }
    .stDownloadButton > button { background: var(--primary-dark) !important; color: white !important; }
    .stDownloadButton > button:hover { background: var(--primary) !important; }
    [data-testid="metric-container"] { background: var(--bg-card) !important; border: 1px solid var(--border) !important; border-radius: 12px !important; padding: 16px !important; }
    [data-testid="metric-container"] label { color: var(--text-muted) !important; }
    [data-testid="metric-container"] [data-testid="stMetricValue"] { color: var(--accent) !important; }
    .stTabs [data-baseweb="tab-list"] { background: var(--bg-card); border-radius: 12px; padding: 4px; gap: 4px; border-bottom: 1px solid var(--border); }
    .stTabs [data-baseweb="tab"] { background: transparent; border-radius: 9px; color: var(--text-muted); padding: 8px 20px; font-size: 0.88rem; }
    .stTabs [aria-selected="true"] { background: var(--primary-light) !important; color: var(--primary-dark) !important; }
    .stRadio > div { gap: 12px; }
    .stRadio [data-baseweb="radio"] span { color: var(--text); }
    div[data-testid="stAlert"] { border-radius: 10px; }
    .stDataFrame { border-radius: 10px; overflow: hidden; }
    [data-baseweb="select"] > div { background: var(--bg-card) !important; border: 1px solid var(--border) !important; border-radius: 10px !important; }
    .stProgress > div > div > div { background: var(--primary) !important; }
    .stDateInput > div > div > input { background: var(--bg-card) !important; border: 1px solid var(--border) !important; border-radius: 10px !important; color: var(--text) !important; }
    .stCheckbox span { color: var(--text); }
    .stMarkdown, .stCaption, label, .stSelectbox label, .stNumberInput label { color: var(--text) !important; }
    </style>
    """

st.markdown(get_css(), unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════
# Hero 头部
# ═══════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
    <div class="hero-title">🎓 校园财务问答助手</div>
    <div class="hero-sub">智能问答 · 精准计算 · 一键生成 · 全程引导 · 实时预警</div>
    <div>
        <span class="badge">💬 AI 问答</span>
        <span class="badge">🧮 自动计算</span>
        <span class="badge">📎 Excel 生成</span>
        <span class="badge">📋 流程引导</span>
        <span class="badge">⚠️ 预算预警</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════
# Session State 初始化
# ═══════════════════════════════════════════════════════════════════════
def _init_state():
    defaults = {
        "chat_history": [],
        "excel_items": [],
        "guide_data": {
            "type": "✈️ 差旅费",
            "purpose": "",
            "dept": "",
            "applicant": "在校本科生",
            "amount": 0.0,
            "items_desc": "",
            "has_invoice": "✅ 有正规发票",
            "invoice_count": 1,
            "special_note": "",
            "approver": "",
            "submit_loc": "",
            "is_urgent": False,
        },
        "guide_generated": False,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()

# ═══════════════════════════════════════════════════════════════════════
# AI 调用（带缓存）
# ═══════════════════════════════════════════════════════════════════════
SYSTEM_PROMPT = """你是本校"校园财务问答助手"，帮助学生、教师、社团成员解答财务问题。

【报销标准参考（通用高校惯例，以我校实际规定为准）】
▸ 差旅费
· 教职工住宿：省内≤350元/天，省外≤500元/天
· 学生住宿：≤200元/天
· 餐补：教职工100元/天，学生50元/天
· 交通：凭正规票据实报实销；打车需说明事由
· 火车：原则购二等座；飞机须提前审批

▸ 社团/班级活动
· ≤500元：辅导员/指导老师签字
· 500~2000元：学院审批
· >2000元：学生处/团委审批
· 必需材料：活动方案、参与名单、正规发票

▸ 办公用品/耗材
· 单价≤500元：直接报销（附发票）
· 500~5000元：需填写固定资产登记单
· >5000元：申请政府/学校集中采购

▸ 科研经费
· 区分横向（企业）与纵向（政府/基金会）课题
· 劳务费须签协议并缴税
· 超预算须提前申请调整

▸ 竞赛/培训
· 需提前填写外出审批表
· 报名费凭收据/发票报销，≤500元指导老师审批，>500元学院审批

▸ 助学贷款
· 每年9月办理；流程：申请→学院审核→资助中心审批→银行放款
· 续贷可在线申请，保持良好信用记录

▸ 学校财务基本原则：统一入账、严禁私收费、全面预算管理、年终公开决算

【回答要求】简洁友好、条理清晰。涉及具体金额注明"以我校实际规定为准"。不确定时建议咨询财务处或查阅校园官网。"""

@cache_data(ttl=300)
def call_ai_cached(messages_json: str):
    messages = json.loads(messages_json)
    resp = client.chat.completions.create(
        model="deepseek-chat",
        messages=messages,
        temperature=0.6,
        max_tokens=1024,
    )
    return resp.choices[0].message.content

def call_ai(messages):
    return call_ai_cached(json.dumps(messages, ensure_ascii=False))

# ═══════════════════════════════════════════════════════════════════════
# 辅助函数（带缓存）
# ═══════════════════════════════════════════════════════════════════════
@cache_data
def parse_reimbursement_cached(text: str):
    d = {}
    m = re.search(r'(\d+(?:\.\d+)?)\s*[天日]', text)
    if m:
        d['days'] = float(m.group(1))
    for pat in [
        r'(?:住宿|酒店|宾馆|旅馆|客栈)[费用]?\s*[：:￥¥]?\s*(\d+(?:\.\d+)?)',
        r'(\d+(?:\.\d+)?)\s*元?[/每]\s*[天日晚夜]',
        r'(\d+(?:\.\d+)?)\s*[元块]/天',
    ]:
        m = re.search(pat, text)
        if m:
            d['hotel_per_day'] = float(m.group(1))
            break
    m = re.search(r'(?:餐[补费]|伙食|用餐|饭补|早餐|午餐|晚餐)[费]?\s*[：:￥¥]?\s*(\d+(?:\.\d+)?)', text)
    if m:
        d['meal_per_day'] = float(m.group(1))
    m = re.search(r'(?:交通|火车|高铁|飞机|机票|大巴|汽车|打车|出租|滴滴)[费票]?\s*[：:￥¥]?\s*(\d+(?:\.\d+)?)', text)
    if m:
        d['transport'] = float(m.group(1))
    m = re.search(r'(?:其他|杂费|额外|补贴)[费]?\s*[：:￥¥]?\s*(\d+(?:\.\d+)?)', text)
    if m:
        d['other'] = float(m.group(1))
    return d if d else None

@cache_data
def parse_items_nl_cached(text: str):
    items = []
    for part in re.split(r'[，,；;\n]+', text):
        part = part.strip()
        if not part:
            continue
        m = re.search(r'^(.+?)\s*[：:￥¥]?\s*(\d+(?:\.\d+)?)\s*[元块]?$', part)
        if m:
            name = m.group(1).strip().strip('：: ')
            price = float(m.group(2))
            if name and price > 0:
                items.append({"物品/费用名称": name, "金额（元）": price})
                continue
        m = re.search(r'^[¥￥]?\s*(\d+(?:\.\d+)?)\s*[元块]?\s+(.+)$', part)
        if m:
            price = float(m.group(1))
            name = m.group(2).strip()
            if name and price > 0:
                items.append({"物品/费用名称": name, "金额（元）": price})
    return items

@cache_data
def build_excel_cached(items_tuple):
    items_list = [{"物品/费用名称": name, "金额（元）": amt} for name, amt in items_tuple]
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    wb = Workbook()
    ws = wb.active
    ws.title = "报销清单"
    thin = Side(style='thin', color="CCCCCC")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    gold_fill = PatternFill("solid", fgColor="D4AF37")
    navy_fill = PatternFill("solid", fgColor="1E3A5F")
    white_fill = PatternFill("solid", fgColor="FFFFFF")
    gray_fill = PatternFill("solid", fgColor="EBF0FA")
    ws.merge_cells('A1:E1')
    ws['A1'] = "报 销 费 用 清 单"
    ws['A1'].font = Font(name='微软雅黑', size=17, bold=True, color="1E3A5F")
    ws['A1'].fill = gold_fill
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40
    ws.merge_cells('A2:E2')
    ws['A2'] = f"生成时间：{datetime.now().strftime('%Y年%m月%d日 %H:%M')}"
    ws['A2'].font = Font(name='微软雅黑', size=9, color="888888")
    ws['A2'].alignment = Alignment(horizontal='right', vertical='center')
    ws.row_dimensions[2].height = 18
    headers = ["序号", "物品 / 费用名称", "单价（元）", "数量", "金额（元）"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(3, col, h)
        c.fill = navy_fill
        c.font = Font(name='微软雅黑', bold=True, color="FFFFFF", size=11)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = bdr
    ws.row_dimensions[3].height = 28
    for i, item in enumerate(items_list):
        r = i + 4
        vals = [i + 1, item["物品/费用名称"], item["金额（元）"], 1, item["金额（元）"]]
        fill = white_fill if i % 2 == 0 else gray_fill
        for col, val in enumerate(vals, 1):
            c = ws.cell(r, col, val)
            c.fill = fill
            c.font = Font(name='微软雅黑', size=10)
            c.border = bdr
            c.alignment = Alignment(horizontal='center' if col in [1,3,4,5] else 'left', vertical='center')
        ws.row_dimensions[r].height = 22
    tr = len(items_list) + 4
    total = sum(x["金额（元）"] for x in items_list)
    ws.merge_cells(f'A{tr}:D{tr}')
    ws[f'A{tr}'] = "合　　计"
    ws[f'A{tr}'].font = Font(name='微软雅黑', bold=True, color="FFFFFF", size=12)
    ws[f'A{tr}'].fill = navy_fill
    ws[f'A{tr}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'E{tr}'] = total
    ws[f'E{tr}'].font = Font(name='微软雅黑', bold=True, color="D4AF37", size=13)
    ws[f'E{tr}'].fill = navy_fill
    ws[f'E{tr}'].alignment = Alignment(horizontal='center', vertical='center')
    for col in range(1, 6):
        ws.cell(tr, col).border = bdr
    ws.row_dimensions[tr].height = 32
    sr = tr + 1
    ws.merge_cells(f'A{sr}:E{sr}')
    ws[f'A{sr}'] = f"报销总金额：¥ {total:.2f} 元"
    ws[f'A{sr}'].font = Font(name='微软雅黑', size=10, color="555555", italic=True)
    ws[f'A{sr}'].alignment = Alignment(horizontal='right', vertical='center')
    ws.row_dimensions[sr].height = 22
    sign_r = tr + 3
    ws.merge_cells(f'A{sign_r}:E{sign_r}')
    ws[f'A{sign_r}'] = "报销人：_______________ 审批人：_______________ 财务审核：_______________ 日期：_______________"
    ws[f'A{sign_r}'].font = Font(name='微软雅黑', size=10, color="444444")
    ws[f'A{sign_r}'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[sign_r].height = 32
    for col, w in [('A',7),('B',30),('C',14),('D',8),('E',14)]:
        ws.column_dimensions[col].width = w
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue(), total

def parse_reimbursement(text):
    return parse_reimbursement_cached(text)

def parse_items_nl(text):
    return parse_items_nl_cached(text)

def build_excel(items_list):
    items_tuple = tuple((item["物品/费用名称"], item["金额（元）"]) for item in items_list)
    return build_excel_cached(items_tuple)

# 流程引导类型库
GUIDE_TYPES = {
    "✈️ 差旅费": {
        "required_docs": ["出差审批单（事先填写）", "正规住宿发票", "交通票据（火车票/机票/汽车票）", "打车费收据（如有）", "报销申请表"],
        "tips": "差旅结束后7个工作日内提交报销，超时可能被拒。",
    },
    "🎉 社团/班级活动费": {
        "required_docs": ["活动申请/方案（提前审批）", "参与人员签到名单", "正规发票或收据", "经费使用明细表", "指导老师/辅导员签字"],
        "tips": "活动经费报销须在活动结束后30天内提交。超500元需学院盖章。",
    },
    "🖨️ 办公用品/耗材": {
        "required_docs": ["采购清单（品名、规格、数量、单价）", "正规增值税发票", "验收确认签字"],
        "tips": "单价超500元须填写固定资产登记表，由资产管理处贴条码后方可领用。",
    },
    "🔬 科研经费": {
        "required_docs": ["项目立项批文或合同复印件", "经费支出申请单", "正规发票（注明课题名称）", "项目负责人签字", "设备类须附验收单"],
        "tips": "科研经费须在预算科目内列支，跨科目需提前申请调整，年底前完成结题报销。",
    },
    "🏆 竞赛/培训费": {
        "required_docs": ["竞赛/培训通知原件", "外出审批表（提前填写）", "报名费发票或收据", "参赛/结业证明（返校后）"],
        "tips": "报名费500元以内指导老师审批；500元以上须学院审批。获奖后可申请额外奖励经费。",
    },
    "📦 其他费用": {
        "required_docs": ["费用说明及用途说明", "正规发票或收据", "经手人签字", "部门负责人审批"],
        "tips": "其他费用需详细说明用途，无法提供发票的须填写无票说明，金额较大时需追加审批。",
    },
}

def get_auto_approver_submit(gtype, amount):
    """根据报销类型和金额返回默认的审批人和提交地点"""
    if "差旅" in gtype:
        if amount <= 500:
            return "指导老师", "学院办公室"
        else:
            return "学院分管院长", "财务处报销大厅"
    elif "社团" in gtype:
        if amount <= 500:
            return "社团指导老师", "团委财务窗口"
        else:
            return "学院学工办", "团委财务窗口"
    elif "办公" in gtype:
        return "实验室/部门负责人", "资产与实验室管理处"
    elif "科研" in gtype:
        return "项目负责人", "科研院财务科"
    elif "竞赛" in gtype:
        return "竞赛指导老师", "教务处实践科"
    else:
        return "部门负责人", "财务处"

def generate_guide_output(data):
    gtype = data["type"]
    type_info = GUIDE_TYPES.get(gtype, GUIDE_TYPES["📦 其他费用"])
    amount = data["amount"]
    has_inv = data["has_invoice"]
    applicant = data["applicant"]
    docs_html = "".join([f'<div class="cl">📄 {doc}</div>' for doc in type_info["required_docs"]])
    warn_html = ""
    if amount > 2000:
        warn_html += '<div class="cl warn">🚨 金额超过2000元，需要学院/学生处级别审批，请提前联系。</div>'
    elif amount > 500:
        warn_html += '<div class="cl info">ℹ️ 金额超过500元，需学院盖章或中级审批，请确认审批流程。</div>'
    else:
        warn_html += '<div class="cl ok">✅ 金额在500元以内，指导老师/辅导员签字即可。</div>'
    if "❌" in has_inv:
        warn_html += '<div class="cl warn">🚨 无票据：请联系商家补开发票，或提交书面无票说明（需部门负责人签字）。</div>'
    elif "⚠️" in has_inv:
        warn_html += '<div class="cl info">ℹ️ 仅有收据：部分情况可接受，建议提前与财务处确认是否可行。</div>'
    else:
        warn_html += '<div class="cl ok">✅ 持有正规发票，票据符合基本要求。</div>'
    if "学生" in applicant and "差旅" in gtype:
        warn_html += '<div class="cl info">ℹ️ 学生住宿标准≤200元/天，超出部分需额外说明。</div>'
    if data.get("special_note"):
        warn_html += f'<div class="cl info">📝 备注：{data["special_note"]}</div>'
    if data.get("is_urgent"):
        warn_html += '<div class="cl warn">⚡ 已标记紧急报销，提交时请在报销单封面注明「加急」并说明原因。</div>'
    
    # 使用自动值（如果用户未填写则使用默认）
    approver = data.get("approver")
    submit_loc = data.get("submit_loc")
    if not approver:
        approver, _ = get_auto_approver_submit(gtype, amount)
    if not submit_loc:
        _, submit_loc = get_auto_approver_submit(gtype, amount)
    
    steps_html = ""
    steps_list = [
        f"整理全部票据并按顺序粘贴在报销单背面",
        f"填写报销申请表（学院财务/行政办公室领取）",
        f"提交给审批人：**{approver}** 签字确认",
        f"携带全部材料前往：**{submit_loc}** 提交",
        f"留存报销单复印件或拍照备份，等待到账（一般3-10个工作日）",
    ]
    for i, s in enumerate(steps_list, 1):
        steps_html += f'<div class="cl">{i}. {s}</div>'
    tips_html = f'<div class="cl info">💡 **小贴士：** {type_info["tips"]}</div>'
    contact_html = '<div class="cl info">📞 如有疑问，请拨打校财务处咨询电话，或前往窗口现场确认。</div>'
    return docs_html + warn_html + steps_html + tips_html + contact_html

# ═══════════════════════════════════════════════════════════════════════
# 主界面 Tabs
# ═══════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "💬 智能问答", "🧮 自动计算", "📎 生成清单", "📋 流程引导", "⚠️ 预算预警"
])

# ──────────────────────────────────────────────────────────────────────
# Tab 1 · 智能问答
# ──────────────────────────────────────────────────────────────────────
with tab1:
    st.markdown("### 💬 智能财务问答")
    st.caption("支持多轮连续对话，所有回答基于学校财务制度及通用高校惯例")
    QUICK_QS = [
        "差旅费住宿标准是多少？", "社团活动报销需要什么材料？",
        "助学贷款怎么办理？", "科研经费报销设备的流程？",
        "报销单丢失怎么办？", "办公用品超过500元怎么处理？",
        "打车费可以报销吗？", "横向课题和纵向课题有什么区别？",
    ]
    st.markdown("**💡 快捷提问（点击直接发送）：**")
    cols = st.columns(4)
    for i, q in enumerate(QUICK_QS):
        with cols[i % 4]:
            if st.button(q, key=f"quick_{i}", use_container_width=True):
                st.session_state.chat_history.append({"role": "user", "content": q})
                st.rerun()
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    if not st.session_state.chat_history:
        st.markdown('<div class="chat-empty">👆 点击快捷提问或在下方输入框开始对话</div>', unsafe_allow_html=True)
    else:
        for msg in st.session_state.chat_history:
            if msg["role"] == "user":
                st.markdown(f'<div class="bubble-label user-label" style="text-align:right;">👤 你</div>'
                            f'<div class="bubble-user">{msg["content"]}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="bubble-label ai-label">🤖 财务助手</div>'
                            f'<div class="bubble-ai">{msg["content"]}</div>', unsafe_allow_html=True)

        if st.session_state.chat_history and st.session_state.chat_history[-1]["role"] == "user":
            with st.spinner("思考中..."):
                msgs = [{"role": "system", "content": SYSTEM_PROMPT}] + st.session_state.chat_history
                reply = call_ai(msgs)
                st.session_state.chat_history.append({"role": "assistant", "content": reply})
                st.rerun()

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([6,1,1])
    with c1:
        user_q = st.text_input("输入问题", key="qa_input", label_visibility="collapsed",
                               placeholder="例：差旅报销需要几天内提交？餐补标准是多少？")
    with c2:
        if st.button("发送 →", key="qa_send", use_container_width=True):
            if user_q.strip():
                st.session_state.chat_history.append({"role": "user", "content": user_q.strip()})
                st.rerun()
    with c3:
        if st.button("清空", key="qa_clear", use_container_width=True):
            st.session_state.chat_history = []
            st.rerun()

# ──────────────────────────────────────────────────────────────────────
# Tab 2 · 自动计算
# ──────────────────────────────────────────────────────────────────────
with tab2:
    st.markdown("### 🧮 自动计算报销金额")
    mode_calc = st.radio("输入方式", ["📝 自然语言描述", "🔢 手动填写表单"], horizontal=True, key="calc_mode")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    if mode_calc == "📝 自然语言描述":
        st.caption("💡 直接描述出差情况，支持多种表达方式")
        examples = [
            "出差3天，住宿400元/天，餐补100元/天，交通费200元",
            "去北京开会5天，酒店350元每晚，伙食补贴80元/天，高铁来回共600元",
            "差旅2日，宾馆每天280，打车费共150",
        ]
        st.markdown(f"**参考格式：** `{examples[0]}` · `{examples[1]}`")
        nl_input = st.text_area("描述出差/报销情况", height=100, key="calc_nl", placeholder=examples[0])
        if st.button("🧮 立即计算", key="calc_nl_btn"):
            with st.spinner("计算中..."):
                result = parse_reimbursement(nl_input)
                if not result:
                    st.warning("⚠️ 未能识别有效数据。请包含「天数」「住宿」「餐补」「交通」等关键词。")
                else:
                    days = result.get('days', 0)
                    hotel = result.get('hotel_per_day', 0)
                    meal = result.get('meal_per_day', 0)
                    trans = result.get('transport', 0)
                    other = result.get('other', 0)
                    hotel_t = days * hotel
                    meal_t = days * meal
                    total = hotel_t + meal_t + trans + other
                    st.success(f"✅ 预计可报销总额：**{total:.2f} 元**")
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("出差天数", f"{days} 天")
                    c2.metric("住宿小计", f"{hotel_t:.0f} 元")
                    c3.metric("餐补小计", f"{meal_t:.0f} 元")
                    c4.metric("交通 + 其他", f"{trans + other:.0f} 元")
                    st.caption("*本结果仅供参考，最终以学校财务处核定为准。*")
    else:
        @st.fragment
        def manual_calc_fragment():
            st.caption("逐项填写各项费用，自动汇总计算")
            c1, c2 = st.columns(2)
            with c1:
                days = st.number_input("出差天数（天）", min_value=0, step=1, key="m_days_frag")
                hotel_day = st.number_input("住宿费（元/天）", min_value=0.0, step=50.0, key="m_hotel_frag")
                meal_day = st.number_input("餐补（元/天）", min_value=0.0, step=10.0, key="m_meal_frag")
            with c2:
                transport = st.number_input("交通费（元，合计）", min_value=0.0, step=10.0, key="m_trans_frag")
                other = st.number_input("其他费用（元）", min_value=0.0, step=10.0, key="m_other_frag")
                person = st.selectbox("申请人身份", ["在校学生", "教职工", "研究生"], key="m_person_frag")
            if st.button("🧮 计算合计", key="calc_manual_frag"):
                hotel_t = days * hotel_day
                meal_t = days * meal_day
                total = hotel_t + meal_t + transport + other
                if person == "在校学生" and hotel_day > 200:
                    st.warning(f"⚠️ 住宿费 {hotel_day:.0f} 元/天超过学生标准（≤200元），超出部分可能无法全额报销。")
                if person == "教职工" and hotel_day > 350:
                    st.warning(f"⚠️ 住宿费 {hotel_day:.0f} 元/天超过省内教职工标准（≤350元）。")
                st.success(f"✅ 预计可报销总额：**{total:.2f} 元**")
                col1, col2, col3, col4, col5 = st.columns(5)
                col1.metric("天数", f"{days}天")
                col2.metric("住宿小计", f"{hotel_t:.0f} 元")
                col3.metric("餐补小计", f"{meal_t:.0f} 元")
                col4.metric("交通费", f"{transport:.0f} 元")
                col5.metric("汇总总额", f"{total:.2f} 元")
                if total > 0:
                    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
                    st.markdown("**费用结构：**")
                    for label, val in [("住宿", hotel_t), ("餐补", meal_t), ("交通", transport), ("其他", other)]:
                        if val > 0:
                            pct = val / total
                            st.write(f"{label} {val:.2f} 元 （{pct*100:.1f}%）")
                            st.progress(pct)
                    st.caption("*本结果仅供参考，最终以学校财务处核定为准。*")
        manual_calc_fragment()

# ──────────────────────────────────────────────────────────────────────
# Tab 3 · 生成清单（自然语言批量输入 + 统一预览）
# ──────────────────────────────────────────────────────────────────────
with tab3:
    st.markdown("### 📎 生成报销清单 Excel")
    mode_xl = st.radio("录入方式", ["📝 自然语言批量输入", "➕ 逐条手动添加"], horizontal=True, key="xl_mode")
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    if mode_xl == "📝 自然语言批量输入":
        st.caption("💡 用逗号或换行分隔每一项，格式非常灵活")
        examples_xl = "打印费50元，笔记本30，马克笔20元\n会场布置费200，茶水服务费80，海报印刷60元"
        nl_xl = st.text_area("输入费用明细", value="", height=140, key="xl_nl", placeholder=examples_xl)
        if st.button("📊 解析并预览", key="xl_parse_btn"):
            with st.spinner("识别中..."):
                items = parse_items_nl(nl_xl)
                if not items:
                    st.warning("⚠️ 未能识别费用条目。请确保每条包含「名称」和「金额」，如：打印费50元")
                else:
                    st.session_state.excel_items = items
                    st.success(f"✅ 成功识别 {len(items)} 条费用")
    else:
        @st.fragment
        def manual_add_fragment():
            st.caption("依次输入每项费用名称和金额，支持随时删除")
            ca, cb, cc = st.columns([3,2,1])
            with ca:
                new_name = st.text_input("费用名称", key="xl_name_frag", placeholder="例：打印费")
            with cb:
                new_amt = st.number_input("金额（元）", min_value=0.0, step=1.0, key="xl_amt_frag")
            with cc:
                st.write("")
                st.write("")
                if st.button("➕ 添加", key="xl_add_btn_frag", use_container_width=True):
                    if new_name.strip() and new_amt > 0:
                        st.session_state.excel_items.append({"物品/费用名称": new_name.strip(), "金额（元）": new_amt})
        manual_add_fragment()

    # 统一展示清单（无论哪种录入方式，只要 excel_items 非空）
    if st.session_state.excel_items:
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown(f"**📋 当前清单（共 {len(st.session_state.excel_items)} 条）**")
        for i, item in enumerate(st.session_state.excel_items):
            col_n, col_p, col_d = st.columns([5,2,1])
            col_n.write(f"**{i+1}.** {item['物品/费用名称']}")
            col_p.write(f"{item['金额（元）']:.2f} 元")
            if col_d.button("🗑️", key=f"del_{i}_main", help="删除此项"):
                st.session_state.excel_items.pop(i)
                st.rerun()
        total_preview = sum(x["金额（元）"] for x in st.session_state.excel_items)
        st.markdown(f"<div style='text-align:right;color:#D4AF37;font-size:1.05rem;margin-top:10px;'>合计：{total_preview:.2f} 元</div>", unsafe_allow_html=True)
        c_dl, c_clr = st.columns([3,1])
        with c_dl:
            excel_data, _ = build_excel(st.session_state.excel_items)
            st.download_button(
                label="📥 下载 Excel 报销清单",
                data=excel_data,
                file_name=f"报销清单_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with c_clr:
            if st.button("清空全部", key="xl_clear_main", use_container_width=True):
                st.session_state.excel_items = []
                st.rerun()
    else:
        st.info("📂 暂无费用条目，请通过上方方式添加。")

# ──────────────────────────────────────────────────────────────────────
# Tab 4 · 流程引导（自动填充审批人和提交地点）
# ──────────────────────────────────────────────────────────────────────
with tab4:
    st.markdown("### 📋 报销流程引导")
    st.caption("填写以下所有信息，点击「重新生成材料清单」即可获得完整报销指引。所有字段随时可修改。")

    @st.fragment
    def guide_fragment():
        # 报销类型
        guide_type = st.selectbox("报销类型", list(GUIDE_TYPES.keys()), index=list(GUIDE_TYPES.keys()).index(st.session_state.guide_data["type"]), key="guide_type_frag")
        st.session_state.guide_data["type"] = guide_type
        amount = st.session_state.guide_data["amount"]

        # 自动获取默认审批人和提交地点（用于提示和预填充）
        default_approver, default_submit_loc = get_auto_approver_submit(guide_type, amount)

        col1, col2 = st.columns(2)
        with col1:
            purpose = st.text_input("报销事由/活动名称", value=st.session_state.guide_data["purpose"], key="guide_purpose_frag")
            st.session_state.guide_data["purpose"] = purpose
            applicant = st.selectbox("申请人身份", ["在校本科生", "在校研究生", "教职工", "博士生"], index=["在校本科生","在校研究生","教职工","博士生"].index(st.session_state.guide_data["applicant"]), key="guide_applicant_frag")
            st.session_state.guide_data["applicant"] = applicant
            amount_input = st.number_input("报销总金额（元）", min_value=0.0, step=10.0, value=st.session_state.guide_data["amount"], key="guide_amount_frag")
            st.session_state.guide_data["amount"] = amount_input
        with col2:
            dept = st.text_input("所在学院/部门", value=st.session_state.guide_data["dept"], key="guide_dept_frag")
            st.session_state.guide_data["dept"] = dept
            has_invoice = st.radio("票据情况", ["✅ 有正规发票", "⚠️ 只有收据", "❌ 无任何票据"], index=["✅ 有正规发票","⚠️ 只有收据","❌ 无任何票据"].index(st.session_state.guide_data["has_invoice"]), key="guide_has_invoice_frag")
            st.session_state.guide_data["has_invoice"] = has_invoice
            invoice_count = st.number_input("发票/收据张数", min_value=0, step=1, value=st.session_state.guide_data["invoice_count"], key="guide_invoice_count_frag")
            st.session_state.guide_data["invoice_count"] = invoice_count
        items_desc = st.text_area("费用明细说明（选填）", value=st.session_state.guide_data["items_desc"], height=80, key="guide_items_desc_frag")
        st.session_state.guide_data["items_desc"] = items_desc
        special_note = st.text_area("特殊说明（选填）", value=st.session_state.guide_data["special_note"], height=60, key="guide_special_note_frag")
        st.session_state.guide_data["special_note"] = special_note
        
        # 审批人和提交地点，预填充默认值，但允许用户修改
        approver = st.text_input("审批人（姓名或职务）", value=st.session_state.guide_data["approver"] if st.session_state.guide_data["approver"] else default_approver, key="guide_approver_frag")
        st.session_state.guide_data["approver"] = approver
        submit_loc = st.text_input("提交地点/窗口", value=st.session_state.guide_data["submit_loc"] if st.session_state.guide_data["submit_loc"] else default_submit_loc, key="guide_submit_loc_frag")
        st.session_state.guide_data["submit_loc"] = submit_loc
        is_urgent = st.checkbox("⚡ 紧急报销（需加急处理）", value=st.session_state.guide_data["is_urgent"], key="guide_is_urgent_frag")
        st.session_state.guide_data["is_urgent"] = is_urgent

        if st.button("📋 重新生成材料清单", key="guide_regenerate_frag", use_container_width=True):
            with st.spinner("生成中..."):
                st.session_state.guide_generated = True
        if st.session_state.guide_generated:
            st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
            st.success("🎉 材料清单已生成！")
            output_html = generate_guide_output(st.session_state.guide_data)
            st.markdown(output_html, unsafe_allow_html=True)
        else:
            st.info("👆 填写完信息后，点击「重新生成材料清单」按钮即可获得详细指引。")
    guide_fragment()

# ──────────────────────────────────────────────────────────────────────
# Tab 5 · 预算预警（简化版：单项目仪表盘 + AI分析 + 合规提醒）
# ──────────────────────────────────────────────────────────────────────
with tab5:
    st.markdown("### ⚠️ 预算预警与智能分析")
    st.caption("设置预算总额、已使用金额和截止日期，系统将自动评估风险并给出AI建议。")

    # 简单输入区域（单个项目）
    col1, col2, col3 = st.columns(3)
    with col1:
        total_budget = st.number_input("预算总额（元）", min_value=0.0, step=100.0, value=10000.0, key="simple_total")
    with col2:
        used_budget = st.number_input("已使用金额（元）", min_value=0.0, step=10.0, value=3500.0, key="simple_used")
    with col3:
        deadline = st.date_input("预算截止日期", value=date(date.today().year, 12, 31), key="simple_deadline")
    project_name = st.text_input("预算项目名称（选填）", value="我的项目", key="simple_name")

    if st.button("📊 生成预警报告", key="simple_btn", use_container_width=True):
        with st.spinner("分析中..."):
            total = float(total_budget)
            used = float(used_budget)
            remaining = total - used
            pct_used = (used / total * 100) if total > 0 else 0
            days_left = (deadline - date.today()).days
            days_total = (deadline - date(date.today().year, 1, 1)).days
            days_passed = max(days_total - days_left, 0)

            # 指标卡片
            st.markdown("#### 📊 预算执行概况")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("预算总额", f"{total:.0f} 元")
            col2.metric("已使用", f"{used:.0f} 元", delta=f"{pct_used:.1f}%")
            col3.metric("剩余预算", f"{remaining:.2f} 元", 
                        delta="充足" if remaining > total*0.3 else ("偏低" if remaining>0 else "超支"),
                        delta_color="normal" if remaining>0 else "inverse")
            col4.metric("距截止日期", f"{max(days_left,0)} 天", 
                        delta="时间充裕" if days_left>30 else ("即将到期" if days_left>0 else "已过期"),
                        delta_color="normal" if days_left>14 else "inverse")

            # 仪表盘（速度表）
            color = "#27AE60" if pct_used < 60 else ("#F39C12" if pct_used < 85 else "#E74C3C")
            fig = go.Figure(go.Indicator(
                mode="gauge+number+delta",
                value=pct_used,
                number={"suffix": "%", "font": {"color": "#D4AF37", "size": 32}},
                delta={"reference": 60, "valueformat": ".1f", "increasing": {"color": "#E74C3C"}, "decreasing": {"color": "#27AE60"}},
                gauge={
                    "axis": {"range": [0,100], "tickcolor": "#4A6A8F", "tickfont": {"color": "#4A6A8F"}},
                    "bar": {"color": color},
                    "bgcolor": "#0C1D35",
                    "bordercolor": "#1A3A5F",
                    "steps": [
                        {"range": [0,60], "color": "rgba(39,174,96,0.1)"},
                        {"range": [60,85], "color": "rgba(243,156,18,0.1)"},
                        {"range": [85,100], "color": "rgba(231,76,60,0.1)"},
                    ],
                    "threshold": {"line": {"color": "#D4AF37", "width": 3}, "value": 80},
                },
                title={"text": f"{project_name}<br>预算使用率", "font": {"color": "#AABDD6", "size": 13}}
            ))
            fig.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font={"family": "Noto Sans SC"}, height=270, margin=dict(t=40,b=0,l=20,r=20))
            st.plotly_chart(fig, use_container_width=True)

            # 进度条（时间 vs 预算）
            time_pct = min((days_passed / max(days_total, 1)) * 100, 100)
            st.markdown("**📅 时间 vs 预算进度对比**")
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(x=[time_pct], y=["时间"], orientation='h', marker_color="#3498DB", name="已过时间", text=[f"{time_pct:.1f}%"], textposition='inside'))
            fig2.add_trace(go.Bar(x=[pct_used], y=["预算"], orientation='h', marker_color=color, name="已用预算", text=[f"{pct_used:.1f}%"], textposition='inside'))
            fig2.update_layout(barmode='overlay', paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font={"color": "#AABDD6"}, xaxis={"range": [0,100], "ticksuffix": "%", "gridcolor": "#1A3050"}, height=150, margin=dict(t=10,b=10,l=60,r=20))
            st.plotly_chart(fig2, use_container_width=True)

            # 简单预测
            if days_passed > 0 and used > 0:
                burn_rate = used / days_passed
                if days_left > 0:
                    projected_total = used + burn_rate * days_left
                    st.info(f"📈 **趋势预测**：当前日均消耗 {burn_rate:.1f} 元。若保持此速度，到期预计总花费 **{projected_total:.0f} 元**，将{'超支' if projected_total > total else '未超支'} {abs(projected_total - total):.0f} 元。")

            # AI 智能分析
            st.markdown("---")
            col_ai, _ = st.columns([1,3])
            with col_ai:
                if st.button("🤖 AI 智能分析建议", key="ai_advice_btn"):
                    with st.spinner("AI 正在分析预算状况..."):
                        days_left_actual = max(days_left, 0)
                        burn_rate_calc = used / max(days_passed, 1)
                        projected_total_calc = used + burn_rate_calc * days_left_actual
                        analysis_prompt = f"""
你是高校财务专家。请根据以下数据给出简短、实用的建议（不超过3句话）：
项目：{project_name}
预算总额：{total} 元
已使用：{used} 元
使用率：{pct_used:.1f}%
剩余天数：{days_left_actual} 天
日均消耗：{burn_rate_calc:.1f} 元/天
到期预测总花费：{projected_total_calc:.0f} 元
请回答：
1. 是否存在超支风险？
2. 给出 1-2 条具体调整建议（如控制支出、提前报销、申请预算调整等）。
要求：简洁、直接、有用。
"""
                        try:
                            response = client.chat.completions.create(
                                model="deepseek-chat",
                                messages=[{"role": "user", "content": analysis_prompt}],
                                temperature=0.5,
                                max_tokens=250,
                            )
                            ai_response = response.choices[0].message.content
                            st.success(f"💡 **AI 建议**：{ai_response}")
                        except Exception as e:
                            st.error(f"AI 分析失败：{e}")

            # 合规性提醒
            st.markdown("**⚖️ 财务合规提醒**")
            if pct_used > 80 and days_left < 30:
                st.warning("⚠️ 预算使用率超过80%且剩余时间不足30天，请尽快处理已发生费用的报销，避免逾期失效。")
            elif pct_used > 80:
                st.warning("⚠️ 预算使用率超过80%，请严格控制后续支出。如确有需要，请提前申请预算调整。")
            elif days_left <= 7 and remaining > 0:
                st.warning(f"⏰ 距截止日期仅剩 {days_left} 天，请抓紧时间完成报销手续。")
            elif remaining < 0:
                st.error("🚨 预算已超支！请立即联系财务处申请追加预算或调整支出计划。")
            else:
                st.success("✅ 当前预算执行平稳，无重大合规风险。请继续保持合规报销。")

# ═══════════════════════════════════════════════════════════════════════
# 页脚
# ═══════════════════════════════════════════════════════════════════════
st.markdown("""
<div style="text-align:center;color:#2A4060;font-size:0.75rem;margin-top:40px;padding:20px 0 10px;">
校园财务问答助手 · 财务问题，秒速搞定
</div>
""", unsafe_allow_html=True)