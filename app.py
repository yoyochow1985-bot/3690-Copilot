"""
HR个性化PD生成器 - Demo版
运行方式: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import json
import re
from pathlib import Path
from docx import Document
import anthropic
import os

# 页面配置
st.set_page_config(
    page_title="HR个性化PD生成器",
    page_icon="📋",
    layout="wide"
)

# 初始化session state
if 'competency_data' not in st.session_state:
    st.session_state.competency_data = None
if 'kpi_template' not in st.session_state:
    st.session_state.kpi_template = None
if 'candidate_data' not in st.session_state:
    st.session_state.candidate_data = None
if 'generated_pd' not in st.session_state:
    st.session_state.generated_pd = None

class CompetencyModelParser:
    """胜任力模型解析器"""

    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.df = None

    def auto_detect_structure(self):
        """自动检测Excel结构"""
        xl_file = pd.ExcelFile(self.excel_path)

        result = {
            "status": "success",
            "sheets": xl_file.sheet_names,
            "warnings": [],
            "total_dimensions": 0
        }

        # 找主sheet
        main_sheet = self._find_main_sheet(xl_file.sheet_names)
        if not main_sheet:
            result["status"] = "error"
            result["warnings"].append("未找到胜任力模型sheet")
            return result

        # 读取数据
        df = pd.read_excel(self.excel_path, sheet_name=main_sheet)
        self.df = df

        # 识别列
        column_mapping = self._detect_columns(df.columns)

        if not column_mapping.get("dimension_col"):
            result["status"] = "error"
            result["warnings"].append("未找到'维度'列")
            return result

        # 统计维度数
        dimension_col = column_mapping["dimension_col"]
        total_dimensions = df[dimension_col].dropna().nunique()

        result["detected_structure"] = column_mapping
        result["total_dimensions"] = total_dimensions
        result["main_sheet"] = main_sheet
        result["sample_dimensions"] = df[dimension_col].dropna().head(5).tolist()

        return result

    def _find_main_sheet(self, sheet_names):
        keywords = ["胜任力", "能力", "模型", "competency"]
        for sheet in sheet_names:
            for keyword in keywords:
                if keyword in str(sheet):
                    return sheet
        return sheet_names[0] if sheet_names else None

    def _detect_columns(self, columns):
        mapping = {
            "dimension_col": None,
            "expected_level_col": None,
            "level_cols": [],
            "module_col": None,
        }

        for col in columns:
            col_str = str(col).lower()

            if any(kw in col_str for kw in ["维度", "dimension"]):
                mapping["dimension_col"] = col
            elif any(kw in col_str for kw in ["期望", "目标", "expected"]):
                mapping["expected_level_col"] = col
            elif any(kw in col_str for kw in ["模块", "module"]):
                mapping["module_col"] = col
            elif re.search(r'[1-5]分|level.*[1-5]', col_str):
                mapping["level_cols"].append(col)

        return mapping

    def extract_competencies(self):
        """提取胜任力数据"""
        result = self.auto_detect_structure()
        if result["status"] != "success":
            return None

        mapping = result["detected_structure"]
        df = self.df

        competencies = []
        dimension_col = mapping["dimension_col"]

        for _, row in df.iterrows():
            dimension = row.get(dimension_col)
            if pd.isna(dimension):
                continue

            competency = {
                "dimension": str(dimension),
                "module": str(row.get(mapping["module_col"])) if mapping["module_col"] else None,
                "expected_level": 3,  # 默认
                "level_descriptions": {}
            }

            # 提取Level描述
            for level_col in mapping["level_cols"]:
                level_num = self._extract_level_number(str(level_col))
                desc = row.get(level_col)
                if not pd.isna(desc):
                    competency["level_descriptions"][level_num] = str(desc)

            competencies.append(competency)

        return competencies

    def _extract_level_number(self, text):
        match = re.search(r'(\d+)', text)
        return int(match.group(1)) if match else 1


class CandidateDataParser:
    """候选人数据解析器"""

    def __init__(self, file_path):
        self.file_path = file_path

    def parse(self):
        """解析Word或Excel文件"""
        file_ext = Path(self.file_path).suffix.lower()

        if file_ext == '.docx':
            return self._parse_word()
        elif file_ext in ['.xlsx', '.xls']:
            return self._parse_excel()
        else:
            return None

    def _parse_word(self):
        """解析Word文档"""
        doc = Document(self.file_path)

        candidate_data = {
            "name": None,
            "position": None,
            "evaluations": [],
            "strengths": [],
            "weaknesses": []
        }

        # 从表格提取
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip() for cell in row.cells]

                if len(cells) >= 2:
                    if "姓名" in cells[0]:
                        candidate_data["name"] = cells[1]
                    elif "应聘职位" in cells[0] or "职位" in cells[0]:
                        candidate_data["position"] = cells[1]
                    elif "综合评价" in cells[0]:
                        candidate_data["evaluations"].append(cells[-1])

        # 提取优劣势
        full_text = "\n".join(candidate_data["evaluations"])
        candidate_data["strengths"] = self._extract_keywords(full_text, "positive")
        candidate_data["weaknesses"] = self._extract_keywords(full_text, "negative")

        return candidate_data

    def _parse_excel(self):
        """解析Excel文件"""
        df = pd.read_excel(self.file_path)

        candidate_data = {
            "name": None,
            "position": None,
            "evaluations": [],
            "strengths": [],
            "weaknesses": []
        }

        # 简化:假设第一行是姓名,第二行是职位
        if len(df) > 0:
            candidate_data["name"] = str(df.iloc[0, 1]) if len(df.columns) > 1 else None
        if len(df) > 1:
            candidate_data["position"] = str(df.iloc[1, 1]) if len(df.columns) > 1 else None

        return candidate_data

    def _extract_keywords(self, text, keyword_type="positive"):
        """提取关键词"""
        positive_kw = {
            "专业": 0.9, "扎实": 0.9, "优秀": 0.85, "突出": 0.85,
            "沟通": 0.8, "积极": 0.8, "主动": 0.8, "认真": 0.8,
            "技术": 0.75, "创新": 0.75, "学习": 0.75
        }

        negative_kw = {
            "经验": -0.5, "应届": -0.5, "较少": -0.5, "不足": -0.6
        }

        keywords_dict = positive_kw if keyword_type == "positive" else negative_kw

        found = []
        for kw, score in keywords_dict.items():
            if kw in text:
                found.append({
                    "keyword": kw,
                    "score": abs(score),
                    "context": text[:100]
                })

        return found


def generate_pd_with_api(candidate_data, competencies, kpi_template, api_provider, api_key):
    """使用多种AI API生成个性化PD"""

    # 构建prompt
    prompt = f"""
你是中黄教育集团的HRBP助手。基于以下数据生成个性化的30天试用期目标。

【候选人信息】
姓名: {candidate_data['name']}
岗位: {candidate_data['position']}

【面试评价】
{chr(10).join(candidate_data['evaluations'][:2])}

【优势】
{', '.join([s['keyword'] for s in candidate_data['strengths'][:5]])}

【劣势】
{', '.join([w['keyword'] for w in candidate_data['weaknesses']])}

【胜任力模型】(共{len(competencies)}个维度)
前5个维度: {', '.join([c['dimension'] for c in competencies[:5]])}

【3690 KPI结构】
{chr(10).join([f"{i+1}. {kpi['name']} (权重{int(kpi['weight']*100)}%)" for i, kpi in enumerate(kpi_template[:5])])}

请生成一份30天目标,包含:
1. 个性化分析摘要(优势Top 3)
2. 根据3690 KPI结构生成具体目标
3. 每个KPI标注是否个性化调整(用🔵优势/🟡短板/💡建议)

输出格式为Markdown,简洁专业。
"""

    if "Claude" in api_provider:
        # 使用Claude API
        client = anthropic.Anthropic(api_key=api_key)
        message = client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=4000,
            messages=[{"role": "user", "content": prompt}]
        )
        return message.content[0].text

    elif "DeepSeek" in api_provider:
        # 使用DeepSeek API
        import requests
        response = requests.post(
            "https://api.deepseek.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            },
            json={
                "model": "deepseek-chat",
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": 4000
            }
        )
        return response.json()["choices"][0]["message"]["content"]

    elif "通义" in api_provider:
        # 使用通义千问API
        import requests
        response = requests.post(
            "https://dashscope.aliyuncs.com/api/v1/services/aigc/text-generation/generation",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            },
            json={
                "model": "qwen-max",
                "input": {"messages": [{"role": "user", "content": prompt}]},
                "parameters": {"max_tokens": 4000}
            }
        )
        return response.json()["output"]["text"]

    elif "智谱" in api_provider:
        # 使用智谱AI API
        import requests
        response = requests.post(
            "https://open.bigmodel.cn/api/paas/v4/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            },
            json={
                "model": "glm-4",
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": 4000
            }
        )
        return response.json()["choices"][0]["message"]["content"]


def generate_pd_local(candidate_data, competencies, kpi_template):
    """本地规则引擎生成(不使用API) - 增强版"""

    # 分析候选人画像
    has_strong_professional = any(
        kw in s['keyword'] for s in candidate_data['strengths']
        for kw in ['专业', '扎实', '优秀', '学科']
    ) and any(s['score'] > 0.85 for s in candidate_data['strengths'])

    has_tech_advantage = any(
        kw in s['keyword'] for s in candidate_data['strengths']
        for kw in ['技术', '信息', '工具']
    )

    is_fresh_graduate = any(
        kw in w['keyword'] for w in candidate_data['weaknesses']
        for kw in ['经验', '应届', '较少']
    )

    wants_班主任 = any('班主任' in str(eval) for eval in candidate_data['evaluations'])

    # 识别需要重点培养的胜任力维度
    weak_competencies = []
    strong_competencies = []

    for comp in competencies[:10]:  # 分析前10个维度
        dimension = comp['dimension']
        # 简化判断:如果维度名出现在面试评价中,认为有基础
        mentioned = any(dimension[:2] in str(eval) for eval in candidate_data['evaluations'])

        if mentioned:
            strong_competencies.append(dimension)
        else:
            weak_competencies.append(dimension)

    pd_content = f"""# {candidate_data['name']} - 30天试用期目标

## 📊 个性化分析摘要

### 优势维度 Top 3
"""

    for idx, strength in enumerate(candidate_data['strengths'][:3], 1):
        pd_content += f"{idx}. **{strength['keyword']}** ({int(strength['score']*100)}%) - {strength['context'][:30]}...\n"

    pd_content += "\n### ⚠️ 发展维度(需重点培养)\n"

    for comp in weak_competencies[:5]:
        pd_content += f"- **{comp}**: 当前水平较低,需系统培养\n"

    pd_content += f"\n### 💡 面试官建议\n"
    for eval in candidate_data['evaluations'][:2]:
        if '建议' in eval or '拓展' in eval or '发展' in eval:
            suggestion = eval.split('。')[0]
            pd_content += f"- {suggestion}\n"

    pd_content += "\n---\n\n## 🎯 30天目标详解\n\n"

    # 根据3690模板生成KPI,结合个性化
    for idx, kpi in enumerate(kpi_template[:5], 1):
        kpi_name = kpi['name']
        weight = int(kpi['weight']*100)

        pd_content += f"### KPI {idx}: {kpi_name} (权重 {weight}%)\n\n"

        # 根据KPI类型匹配个性化策略
        if '教学设计' in kpi_name or '教案' in kpi_name:
            if has_strong_professional:
                pd_content += "**个性化策略**: 🔵 专业能力强,提高挑战度\n\n"
                pd_content += "| 项目 | 标准目标 | 🔵个性化调整 | 成功标准 | 支持措施 |\n"
                pd_content += "|------|----------|------------|---------|----------|\n"
                pd_content += f"| 教案设计 | 完成1份教案 | **增加要求**:额外完成1份跨学科探究单元设计 | 常规教案≥85分<br>探究单元通过评审 | 提供PYP单元设计模板 |\n"
            else:
                pd_content += "**个性化策略**: ⚪ 维持标准培养\n\n"
                pd_content += "| 项目 | 标准目标 | 成功标准 |\n"
                pd_content += "|------|----------|----------|\n"
                pd_content += f"| 教案设计 | 完成1份教案 | ≥80分 |\n"

        elif '课堂' in kpi_name or '执行' in kpi_name:
            if is_fresh_graduate:
                pd_content += "**个性化策略**: 🟡 应届生经验少,增加辅导支持\n\n"
                pd_content += "| 项目 | 标准目标 | 🟡个性化调整 | 成功标准 | 支持措施 |\n"
                pd_content += "|------|----------|------------|---------|----------|\n"
                pd_content += f"| 观摩学习 | 听评课≥5节 | **分阶段降低难度**:<br>第1-2周:观摩为主<br>第3-4周:试讲+反馈 | 提交5份听课笔记<br>第4周试讲学生参与度≥70% | 师傅每周陪同≥3次 |\n"
                pd_content += f"| 磨课演练 | 试讲2次 | **降低初期标准**:<br>讲授时间≤50%(逐步过渡至30%) | 第2次试讲获得师傅认可 | 录像回看+逐句点评 |\n"
            else:
                pd_content += "**个性化策略**: ⚪ 标准培养\n\n"
                pd_content += "| 项目 | 标准目标 | 成功标准 |\n"
                pd_content += "|------|----------|----------|\n"
                pd_content += f"| 观摩与试讲 | 听评课≥5节,试讲2次 | 试讲≥80分 |\n"

        elif '家校' in kpi_name or '沟通' in kpi_name:
            has_comm_strength = any('沟通' in s['keyword'] for s in candidate_data['strengths'])
            if has_comm_strength:
                pd_content += "**个性化策略**: 🔵 沟通能力好,可承担更高要求\n\n"
                pd_content += "| 项目 | 标准目标 | 🔵个性化调整 | 成功标准 |\n"
                pd_content += "|------|----------|------------|----------|\n"
                pd_content += f"| 家长沟通 | 24h响应率≥85% | **提高标准**:响应率≥90%,并主动反馈学生进步 | 家长满意度≥90% |\n"
            else:
                pd_content += "**个性化策略**: ⚪ 标准培养\n\n"
                pd_content += "| 项目 | 标准目标 | 成功标准 |\n"
                pd_content += "|------|----------|----------|\n"
                pd_content += f"| 家长沟通 | 24h响应率≥85% | 无投诉 |\n"

        elif '评价' in kpi_name or '工具' in kpi_name or '技术' in kpi_name:
            if has_tech_advantage:
                pd_content += "**个性化策略**: 🔵 信息技术能力强,发挥优势\n\n"
                pd_content += "| 项目 | 标准目标 | 🔵个性化调整 | 成功标准 |\n"
                pd_content += "|------|----------|------------|----------|\n"
                pd_content += f"| 数字化工具 | 使用评价工具 | **创新应用**:尝试AI辅助批改/Seewo互动工具 | 使用≥5次,学生反馈正面 |\n"
            else:
                pd_content += "**个性化策略**: ⚪ 标准培养\n\n"
                pd_content += "| 项目 | 标准目标 | 成功标准 |\n"
                pd_content += "|------|----------|----------|\n"
                pd_content += f"| 评价工具 | 规范使用 | 使用≥3次 |\n"

        elif '专业发展' in kpi_name or '培训' in kpi_name:
            if wants_班主任:
                pd_content += "**个性化策略**: 💡 响应成长意愿,增加班主任培养\n\n"
                pd_content += "| 项目 | 标准目标 | 💡个性化调整 | 成功标准 |\n"
                pd_content += "|------|----------|------------|----------|\n"
                pd_content += f"| 专业学习 | 参加常规培训 | **额外增加**:观摩2次班会+跟随班主任见习 | 提交班主任见习报告 |\n"
            else:
                pd_content += "**个性化策略**: ⚪ 标准培养\n\n"
                pd_content += "| 项目 | 标准目标 | 成功标准 |\n"
                pd_content += "|------|----------|----------|\n"
                pd_content += f"| 专业发展 | 参加培训,完成反思 | 提交学习笔记 |\n"

        pd_content += "\n"

    pd_content += """---

## 📋 支持资源

### 指定师傅
- **学科导师**: (待指定)资深地理教师
"""

    if is_fresh_graduate:
        pd_content += "- **辅导频次**: ⚠️ 应届生需高频支持,建议每周≥3次面对面辅导\n"
    else:
        pd_content += "- **辅导频次**: 每周≥2次\n"

    if wants_班主任:
        pd_content += "- **班主任导师**: (待指定)优秀班主任,用于见习观摩\n"

    pd_content += """
### 推荐培训
1. IB理念与教学法(第1周)
2. 探究式单元设计(第3周)
"""

    if has_tech_advantage:
        pd_content += "3. AI工具在教学中的应用(第4周) - 🔵 推荐参加\n"

    pd_content += """
---

**生成方式**: 本地规则引擎(增强版)
**生成时间**: 2026-07-27
**个性化依据**:
- 面试评价关键词提取
- 胜任力模型差距分析
- 3690 KPI结构匹配
"""

    return pd_content


# ========== Streamlit UI ==========

st.title("📋 HR个性化PD生成器 - Demo版")
st.markdown("---")

# 侧边栏:配置
with st.sidebar:
    st.header("⚙️ 配置")

    # API配置
    use_api = st.checkbox("使用AI增强质量", value=False)

    if use_api:
        api_provider = st.selectbox(
            "选择AI服务商",
            ["Claude (Anthropic)", "DeepSeek", "通义千问 (Qwen)", "智谱AI (GLM)"],
            index=0
        )

        api_key = st.text_input(
            "API Key",
            type="password",
            value=os.environ.get("ANTHROPIC_API_KEY", "")
        )

        if not api_key:
            st.warning("请输入API Key")

        # API配置说明
        with st.expander("💡 API获取指南"):
            if "Claude" in api_provider:
                st.markdown("访问 https://console.anthropic.com 获取")
            elif "DeepSeek" in api_provider:
                st.markdown("访问 https://platform.deepseek.com 获取")
            elif "通义" in api_provider:
                st.markdown("访问 https://dashscope.aliyun.com 获取")
            elif "智谱" in api_provider:
                st.markdown("访问 https://open.bigmodel.cn 获取")

    st.markdown("---")
    st.markdown("### 💡 使用说明")
    st.markdown("""
    1. 上传胜任力模型Excel
    2. 上传3690评估表格Excel
    3. 上传候选人面试数据
    4. 点击生成
    5. 下载Markdown文档
    """)

# 主界面
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("📊 步骤1: 胜任力模型")

    competency_file = st.file_uploader(
        "上传胜任力模型Excel",
        type=['xlsx', 'xls'],
        key="competency"
    )

    if competency_file:
        # 保存临时文件
        temp_path = f"temp_competency.xlsx"
        with open(temp_path, "wb") as f:
            f.write(competency_file.getbuffer())

        try:
            parser = CompetencyModelParser(temp_path)
            result = parser.auto_detect_structure()

            if result["status"] == "success":
                st.success(f"✅ 识别到 {result['total_dimensions']} 个胜任力维度")

                with st.expander("查看识别详情"):
                    st.json(result["detected_structure"])
                    st.write("示例维度:", result.get("sample_dimensions", []))

                # 提取数据
                competencies = parser.extract_competencies()
                st.session_state.competency_data = competencies

                st.info(f"已加载 {len(competencies)} 个维度")
            else:
                st.error("❌ " + "\n".join(result["warnings"]))

        except Exception as e:
            st.error(f"解析失败: {str(e)}")

with col2:
    st.subheader("📋 步骤2: 3690评估表格")

    kpi_file = st.file_uploader(
        "上传3690模板Excel",
        type=['xlsx', 'xls'],
        key="kpi_file"
    )

    if kpi_file:
        temp_path = f"temp_kpi.xlsx"
        with open(temp_path, "wb") as f:
            f.write(kpi_file.getbuffer())

        try:
            # 读取3690模板
            df_kpi = pd.read_excel(temp_path, sheet_name="教师考核细则")

            st.success("✅ 3690模板解析成功")

            with st.expander("查看KPI结构"):
                st.write(f"识别到 {len(df_kpi)} 个KPI指标")
                st.dataframe(df_kpi.head(5))

            # 提取KPI结构
            kpi_structure = []
            for idx, row in df_kpi.iterrows():
                if pd.notna(row.get('KPI指标')):
                    # 解析权重(支持百分比格式如"25%"或小数格式如0.25)
                    weight_raw = row.get('权重', 0)
                    if isinstance(weight_raw, str):
                        # 移除%符号并转换
                        weight_raw = weight_raw.replace('%', '').strip()
                        weight = float(weight_raw) / 100 if float(weight_raw) > 1 else float(weight_raw)
                    else:
                        weight = float(weight_raw) if weight_raw > 1 else float(weight_raw)

                    kpi_structure.append({
                        "name": str(row['KPI指标']),
                        "weight": weight,
                        "description": str(row.get('评分标准(0-100分)', ''))
                    })

            st.session_state.kpi_template = kpi_structure
            st.info(f"已加载 {len(kpi_structure)} 个KPI")

        except Exception as e:
            st.error(f"解析失败: {str(e)}")
            st.info("💡 提示: 请确保Excel包含'教师考核细则'sheet,且有'KPI指标'和'权重'列")

with col3:
    st.subheader("👤 步骤3: 候选人数据")

    candidate_files = st.file_uploader(
        "上传候选人相关文件(可多选,最多5个)",
        type=['docx', 'xlsx', 'xls', 'pdf'],
        key="candidate",
        accept_multiple_files=True,
        help="支持:简历、面试评价、听评课记录等"
    )

    if candidate_files:
        if len(candidate_files) > 5:
            st.error("⚠️ 最多上传5个文件,请重新选择")
        else:
            st.success(f"✅ 已选择 {len(candidate_files)} 个文件")

            # 合并解析所有文件
            merged_data = {
                "name": None,
                "position": None,
                "evaluations": [],
                "strengths": [],
                "weaknesses": []
            }

            for idx, file in enumerate(candidate_files, 1):
                file_ext = Path(file.name).suffix
                temp_path = f"temp_candidate_{idx}{file_ext}"

                with open(temp_path, "wb") as f:
                    f.write(file.getbuffer())

                try:
                    parser = CandidateDataParser(temp_path)
                    file_data = parser.parse()

                    if file_data:
                        # 合并数据
                        if not merged_data["name"] and file_data.get("name"):
                            merged_data["name"] = file_data["name"]
                        if not merged_data["position"] and file_data.get("position"):
                            merged_data["position"] = file_data["position"]

                        merged_data["evaluations"].extend(file_data.get("evaluations", []))
                        merged_data["strengths"].extend(file_data.get("strengths", []))
                        merged_data["weaknesses"].extend(file_data.get("weaknesses", []))

                except Exception as e:
                    st.warning(f"文件 {file.name} 解析失败: {str(e)}")

            if merged_data["name"]:
                # 去重优势/劣势关键词
                seen_strengths = set()
                unique_strengths = []
                for s in merged_data["strengths"]:
                    if s['keyword'] not in seen_strengths:
                        seen_strengths.add(s['keyword'])
                        unique_strengths.append(s)
                merged_data["strengths"] = unique_strengths

                seen_weaknesses = set()
                unique_weaknesses = []
                for w in merged_data["weaknesses"]:
                    if w['keyword'] not in seen_weaknesses:
                        seen_weaknesses.add(w['keyword'])
                        unique_weaknesses.append(w)
                merged_data["weaknesses"] = unique_weaknesses

                with st.expander("查看合并后的解析结果"):
                    st.write("**候选人:**", merged_data.get('name'))
                    st.write("**岗位:**", merged_data.get('position'))
                    st.write("**提取的评价数量:**", len(merged_data['evaluations']))
                    st.write("**优势关键词:**", [s['keyword'] for s in merged_data['strengths']])
                    st.write("**劣势关键词:**", [w['keyword'] for w in merged_data['weaknesses']])

                st.session_state.candidate_data = merged_data
            else:
                st.warning("未能从任何文件中提取到有效数据")

st.markdown("---")

# 生成按钮
st.subheader("🚀 步骤4: 生成个性化PD")

col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])

with col_btn1:
    if st.button("🎯 生成PD", type="primary", use_container_width=True):
        if not st.session_state.competency_data:
            st.error("❌ 请先上传胜任力模型")
        elif not st.session_state.kpi_template:
            st.error("❌ 请先上传3690评估表格")
        elif not st.session_state.candidate_data:
            st.error("❌ 请先上传候选人数据")
        else:
            with st.spinner("正在生成..."):
                try:
                    if use_api and api_key:
                        pd_content = generate_pd_with_api(
                            st.session_state.candidate_data,
                            st.session_state.competency_data,
                            st.session_state.kpi_template,
                            api_provider,
                            api_key
                        )
                    else:
                        pd_content = generate_pd_local(
                            st.session_state.candidate_data,
                            st.session_state.competency_data,
                            st.session_state.kpi_template
                        )

                    st.session_state.generated_pd = pd_content
                    st.success("✅ 生成完成!")

                except Exception as e:
                    st.error(f"生成失败: {str(e)}")

with col_btn2:
    if st.button("🔄 重置", use_container_width=True):
        st.session_state.competency_data = None
        st.session_state.kpi_template = None
        st.session_state.candidate_data = None
        st.session_state.generated_pd = None
        st.rerun()

# 显示生成结果
if st.session_state.generated_pd:
    st.markdown("---")
    st.subheader("📄 生成结果")

    # 显示内容
    with st.expander("查看完整内容", expanded=True):
        st.markdown(st.session_state.generated_pd)

    # 下载按钮
    st.download_button(
        label="⬇️ 下载Markdown文件",
        data=st.session_state.generated_pd,
        file_name=f"{st.session_state.candidate_data.get('name', '候选人')}_30-60-90计划.md",
        mime="text/markdown"
    )

    st.info("💡 提示: 下载后可以用Word打开Markdown文件,或使用Pandoc转换为Word格式")

st.markdown("---")
st.caption("Demo版本 v0.1 | 仅供测试使用")
