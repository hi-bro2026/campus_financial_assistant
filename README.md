# 🎓 校园财务问答助手

AI 驱动的校园财务报销辅助系统，帮助学生、班委、社团负责人和教师快速解决财务报销问题。

**🏆 比赛演示版 v2.1**

## 🚀 功能模块

| 模块 | 说明 |
|------|------|
| 🏠 **首页总览** | KPI 动画卡片、报销类型饼图、费用对比表、演示案例一键加载 |
| 💬 **智能问答** | DeepSeek AI 自然语言对话，打字机效果流式输出，6 类报销场景全覆盖 |
| 🧮 **自动计算** | 差旅/活动/办公/科研/竞赛费用自动核算，动画进度条展示费用构成 |
| 📎 **生成清单** | 报销材料清单 Excel 导出，一键生成规范报销单 Word 文档 |
| 📋 **流程引导** | 5 类报销类型库，按步骤检查材料准备情况，AI 生成报销申请说明 |
| ⚠️ **预算预警** | 输入预算与已支出金额，AI 风险分析 + Plotly 仪表盘可视化 |

## 🛠 技术栈

- **前端**: Streamlit、Plotly、纯 CSS/JS 动画
- **AI**: DeepSeek Chat API（流式调用）
- **导出**: openpyxl (Excel)、python-docx (Word)
- **部署**: Streamlit Community Cloud

## 📦 本地运行

```bash
# 1. 克隆仓库
git clone https://github.com/gzmakes/campus_financial_assistant.git
cd campus_financial_assistant

# 2. 安装依赖
pip install -r requirements.txt

# 3. 配置 API Key
mkdir -p .streamlit
echo 'DEEPSEEK_API_KEY = "你的DeepSeek API Key"' > .streamlit/secrets.toml

# 4. 启动
streamlit run app.py
```

## 🌐 在线体验

访问 Streamlit Cloud 公开链接（无需安装）：

➡️ **[campus-financial-assistant.streamlit.app](https://campusfinancialassistant.streamlit.app)**

## 📁 项目结构

```
campus_finance_assistant/
├── app.py              # 主程序（2080+ 行）
├── requirements.txt    # Python 依赖
├── assets/
│   └── logo.jpg        # 校徽图片
└── .streamlit/
    └── secrets.toml    # API Key 配置（不提交 Git）
```

---

<p align="center">
  <sub>数据科学与人工智能学院 · 2025级 · 郭志明</sub>
</p>
