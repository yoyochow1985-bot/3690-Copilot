"""
HR个性化PD生成器 - 黑客松线下Demo版
特点: 预加载数据 + 预生成结果，无需API，适合现场演示
运行方式: streamlit run app_demo_offline.py
"""

import streamlit as st
import pandas as pd
import json
from pathlib import Path
from docx import Document
import time

# 页面配置
st.set_page_config(
    page_title="HR个性化PD生成器 - Demo",
    page_icon="🎯",
    layout="wide"
)

# 预生成的结果内容（基于LXM真实案例）
PREGENERATED_RESULT = """
# 🧬 AI Growth DNA 分析报告

**候选人:** 李晓明 (LXM)
**岗位:** IB PYP数学教师（应届生）
**分析时间:** 2026-07-31

---

## 📊 能力画像总览

### 🔵 核心优势 (Top 3)

1. **学科专业功底扎实** (匹配度: 92%)
   - 数学专业背景 + 教育学辅修
   - 试讲课逻辑清晰、概念准确
   - 建议: 直接承担核心课程设计

2. **学习能力与主动性强** (匹配度: 88%)
   - 自主学习IB PYP理念
   - 主动了解探究式教学法
   - 建议: 加速进入跨学科项目

3. **技术工具应用意识** (匹配度: 75%)
   - 简历提及数字化教学工具
   - 有潜力整合AI辅助工具
   - 建议: 引导探索GeoGebra/Desmos等

### 🟡 发展短板

1. **课堂管理经验不足** (应届生共性)
   - 缺乏真实班级管理经验
   - 需要导师密集支持

2. **家校沟通待培养**
   - 无实际家长沟通案例
   - 需提供模板和演练机会

---

## 🎯 个性化30-60-90天目标

基于3690评估体系 + 个性化调整

### 📅 第1-30天：夯实基础 + 发挥优势

#### 1. 教学设计能力 (权重25%)

**标准目标:**
- 完成4个单元的教学设计
- 参与集体备课3次

**🔵 个性化调整（优势导向）:**
- ✅ **降低设计难度要求**: 由于专业功底强，可直接设计完整单元（不分拆小节）
- ✅ **增加创新空间**: 鼓励融入数字化工具（如GeoGebra互动演示）
- 💡 **导师建议**: 前2周导师审核，第3-4周独立设计

**具体任务:**
- Week 1-2: 在导师指导下设计"数字与运算"单元（包含3课时）
- Week 3: 独立设计"图形与空间"单元，融入1个数字工具演示
- Week 4: 参与年级组跨学科项目设计（数学部分）

---

#### 2. 课堂教学能力 (权重30%)

**标准目标:**
- 完成16节正式课（每周4节）
- 听评课8次
- 导师听课4次

**🟡 个性化调整（短板防护）:**
- ⚠️ **降低初期课堂管理要求**: 前2周由导师陪同上课，观察学生管理技巧
- ⚠️ **增加课堂管理专项培训**: 第1周参加"低龄段课堂规则建立"工作坊
- 💡 **导师建议**: 使用结构化课堂流程（铃声、手势、计时器等辅助工具）

**具体任务:**
- Week 1: 导师陪同上课4节，课后1对1反馈（重点：课堂规则建立）
- Week 2: 半独立上课4节（导师在场但不干预），记录管理问题
- Week 3-4: 独立上课8节，每周导师听课1次，听同组老师课4次

---

#### 3. 学生评价与反馈 (权重20%)

**标准目标:**
- 完成2次形成性评估设计
- 批改作业100份以上
- 给予学生个性化反馈20人次

**🔵 个性化调整（优势延伸）:**
- ✅ **提高评估设计要求**: 由于学科功底强，可设计开放式评估任务
- 💡 **导师建议**: 尝试使用"成长型反馈"框架（What I see / What I wonder）

**具体任务:**
- Week 1-2: 观摩导师批改作业，学习反馈语言（至少50份）
- Week 3: 设计1次开放式评估（如数学日记、实践任务）
- Week 4: 独立完成50份作业批改 + 给10名学生书面成长反馈

---

#### 4. 家校沟通与协作 (权重15%)

**标准目标:**
- 参与家长会1次
- 家校沟通记录5次

**🟡 个性化调整（短板补强）:**
- ⚠️ **增加模板支持**: 提供3种常见沟通场景话术模板
- ⚠️ **降低沟通复杂度**: 第1月仅处理常规沟通（作业反馈、到校情况），问题性沟通由导师主导
- 💡 **导师建议**: 每次正式沟通前与导师预演

**具体任务:**
- Week 1: 学习家校沟通案例3个，了解平台使用
- Week 2-3: 在导师指导下完成3次家长沟通（使用模板）
- Week 4: 独立完成2次常规沟通，参与导师主导的家长会1次

---

#### 5. 专业发展与团队协作 (权重10%)

**标准目标:**
- 参与教研活动4次
- 完成IB PYP理念培训

**🔵 个性化调整（优势强化）:**
- ✅ **提高参与深度**: 由于已有PYP基础，可承担教研分享角色
- 💡 **导师建议**: 分享1次"数字工具在数学探究中的应用"

**具体任务:**
- Week 1: 参与新教师培训 + PYP理念工作坊2次
- Week 2-3: 参加年级组教研活动2次（观察学习）
- Week 4: 在教研组做5分钟分享：我的试讲反思 + 数字工具尝试

---

### 📅 第31-60天：独立实践 + 能力突破

**重点方向:**
1. **课堂管理能力达标** - 能独立管理30人班级40分钟
2. **跨学科项目参与** - 完成1个跨学科单元合作设计
3. **家校沟通信心建立** - 独立处理10次以上家长沟通
4. **教学反思深化** - 建立个人教学日志，每周反思1次

**个性化任务示例（第45天）:**
- 🎯 设计并实施1节融合数字工具的公开课（GeoGebra或Desmos）
- 🎯 与艺术老师合作设计"几何与艺术"跨学科项目
- 🎯 完成15份学生成长档案（portfolio）记录

---

### 📅 第61-90天：独当一面 + 特长展现

**重点方向:**
1. **独立承担完整单元** - 从设计到评估全流程
2. **技术特长应用** - 成为年级组数字化教学先锋
3. **导师角色转变** - 从被支持到能支持（协助新实习生）
4. **转正答辩准备** - 整理教学成果集 + 个人成长档案

**个性化任务示例（第75天）:**
- 🏆 设计并实施1个完整的探究式数学单元（3周课程）
- 🏆 在全校教研会分享"数字工具赋能数学探究"（15分钟）
- 🏆 辅导1名实习生完成试讲准备

---

## 💡 成功关键建议

### 给候选人 LXM:
1. **发挥优势** - 你的学科功底是核心竞争力，大胆设计有挑战性的任务
2. **正视短板** - 课堂管理需要刻意练习，使用辅助工具（计时器、手势、音乐）
3. **主动求助** - 前30天高频找导师反馈（每周至少2次深度对话）

### 给导师:
1. **前2周高强度陪同** - 课堂管理需要现场示范 + 即时反馈
2. **提供结构化工具** - 课堂流程卡片、家校沟通模板、评估设计框架
3. **放手时机** - 第3周起逐步减少干预，允许试错

### 给HR/教学主管:
1. **技术资源支持** - 提供GeoGebra/Desmos培训，鼓励工具创新
2. **跨学科机会** - 第2月起安排参与PYP跨学科项目设计
3. **转正评估侧重点** - 重点看成长曲线（而非绝对水平），尤其课堂管理进步

---

## 📈 预期成果

**30天后:**
- ✅ 能独立设计符合PYP理念的数学单元
- ✅ 能在导师协助下管理30人班级
- ✅ 建立基本家校沟通信心

**60天后:**
- ✅ 能独立上好40分钟课（课堂管理达标）
- ✅ 完成1个跨学科项目合作
- ✅ 熟练使用至少1种数字化教学工具

**90天后:**
- ✅ 能独立承担完整单元（设计-实施-评估）
- ✅ 成为年级组数字化教学先锋
- ✅ 具备辅导实习生的能力
- 🎯 **转正通过率预测: 92%**（基于历史数据：应届生中学科强+主动学习型通过率最高）

---

## 🔍 数据来源

- ✅ 胜任力模型: 中黄IB教师能力框架（12个维度）
- ✅ 3690 KPI: 教师试用期考核细则（5大维度）
- ✅ 候选人数据: 简历 + 面试评价 + 试讲反馈
- 🤖 AI匹配引擎: Claude 3.5 Sonnet
- ⏱️ 生成时间: 28秒

---

*本计划由 HR个性化PD生成器 v0.2 自动生成*
*最后更新: 2026-07-31*
"""

# 初始化session state
if 'demo_mode' not in st.session_state:
    st.session_state.demo_mode = True
if 'step' not in st.session_state:
    st.session_state.step = 0  # 0=初始, 1=已上传文件, 2=已生成
if 'show_result' not in st.session_state:
    st.session_state.show_result = False

# 标题区域
st.title("🎯 HR个性化PD生成器")
st.markdown("### AI驱动的30-60-90天试用期计划定制系统")

# Demo模式提示
st.info("📢 **黑客松Demo模式** | 本版本使用预加载数据，无需API，适合现场快速演示")

st.markdown("---")

# 步骤指示器
col_indicator1, col_indicator2, col_indicator3 = st.columns(3)
with col_indicator1:
    status1 = "✅" if st.session_state.step >= 1 else "1️⃣"
    st.markdown(f"### {status1} 上传文件")
with col_indicator2:
    status2 = "✅" if st.session_state.step >= 2 else "2️⃣"
    st.markdown(f"### {status2} AI分析")
with col_indicator3:
    status3 = "✅" if st.session_state.step >= 2 else "3️⃣"
    st.markdown(f"### {status3} 查看结果")

st.markdown("---")

# 文件上传区域
if st.session_state.step == 0:
    st.subheader("📂 步骤1: 上传必要文件")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("#### 📊 胜任力模型")
        st.file_uploader(
            "上传Excel文件",
            type=['xlsx', 'xls'],
            key="comp_upload",
            help="系统将自动识别胜任力维度"
        )
        st.caption("示例: 中黄IB教师能力框架")

    with col2:
        st.markdown("#### 📋 3690评估模板")
        st.file_uploader(
            "上传Excel文件",
            type=['xlsx', 'xls'],
            key="kpi_upload",
            help="30-60-90天考核指标"
        )
        st.caption("示例: 教师试用期考核细则")

    with col3:
        st.markdown("#### 👤 候选人资料")
        st.file_uploader(
            "上传Word/Excel文件（可多选，最多5个）",
            type=['docx', 'xlsx', 'xls'],
            key="candidate_upload",
            accept_multiple_files=True,
            help="简历、面试评价、试讲反馈等"
        )
        st.caption("示例: 简历.docx + 面试评价.xlsx + 试讲反馈.docx")

    st.markdown("---")

    # 快速演示按钮
    col_demo1, col_demo2, col_demo3 = st.columns([1, 2, 1])
    with col_demo2:
        if st.button("🚀 使用示例数据快速演示", type="primary", use_container_width=True):
            with st.spinner("正在加载示例数据..."):
                time.sleep(1.5)  # 模拟加载
                st.session_state.step = 1
                st.rerun()

# 文件已上传，显示解析结果
elif st.session_state.step == 1:
    st.success("✅ 文件上传成功！正在解析...")

    time.sleep(0.5)  # 短暂延迟增加真实感

    # 显示解析结果
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("#### 📊 胜任力模型")
        st.info("**文件:** temp_competency.xlsx")
        with st.expander("查看解析结果"):
            st.write("✅ 识别到 **12个能力维度**")
            st.write("**核心维度:**")
            st.write("- IB全人观理解")
            st.write("- 探究式教学设计")
            st.write("- PYP课程实施")
            st.write("- 学科专业知识")
            st.write("- 学情分析能力")
            st.write("...")

    with col2:
        st.markdown("#### 📋 3690评估模板")
        st.info("**文件:** temp_kpi.xlsx")
        with st.expander("查看解析结果"):
            st.write("✅ 识别到 **5个KPI维度**")
            st.write("**权重分布:**")
            st.write("- 教学设计能力 (25%)")
            st.write("- 课堂教学能力 (30%)")
            st.write("- 学生评价反馈 (20%)")
            st.write("- 家校沟通协作 (15%)")
            st.write("- 专业发展协作 (10%)")

    with col3:
        st.markdown("#### 👤 候选人资料")
        st.info("**已上传:** 3个文件")
        with st.expander("查看解析结果"):
            st.write("**文件列表:**")
            st.write("- 📄 李晓明_简历.docx")
            st.write("- 📋 面试评价表.xlsx")
            st.write("- 📝 试讲反馈.docx")
            st.write("")
            st.write("**合并解析结果:**")
            st.write("**姓名:** 李晓明 (LXM)")
            st.write("**岗位:** IB PYP数学教师")
            st.write("**优势关键词:** (从3个文件提取)")
            st.write("- 专业 (0.9)")
            st.write("- 扎实 (0.9)")
            st.write("- 主动 (0.8)")
            st.write("- 学习能力强 (0.8)")
            st.write("**短板关键词:**")
            st.write("- 应届生 (0.5)")
            st.write("- 经验不足 (0.6)")

    st.markdown("---")

    # 生成按钮
    col_gen1, col_gen2, col_gen3 = st.columns([1, 2, 1])
    with col_gen2:
        if st.button("🤖 AI智能分析并生成个性化计划", type="primary", use_container_width=True):
            # 显示生成动画
            progress_bar = st.progress(0)
            status_text = st.empty()

            steps = [
                ("解析胜任力模型...", 20),
                ("提取3690 KPI结构...", 40),
                ("分析候选人画像...", 60),
                ("AI匹配差距识别...", 80),
                ("生成个性化计划...", 100)
            ]

            for step_text, progress in steps:
                status_text.text(step_text)
                progress_bar.progress(progress)
                time.sleep(0.6)

            status_text.text("✅ 生成完成!")
            time.sleep(0.5)

            st.session_state.step = 2
            st.session_state.show_result = True
            st.rerun()

# 显示生成结果
elif st.session_state.step == 2:
    st.success("🎉 个性化30-60-90天计划生成完成！")

    # 关键指标卡片
    col_metric1, col_metric2, col_metric3, col_metric4 = st.columns(4)
    with col_metric1:
        st.metric("生成时间", "28秒", delta="-95% vs 人工")
    with col_metric2:
        st.metric("匹配维度", "12个", delta="100%覆盖")
    with col_metric3:
        st.metric("个性化调整", "18处", delta="基于AI分析")
    with col_metric4:
        st.metric("预测通过率", "92%", delta="+15% vs 平均")

    st.markdown("---")

    # 显示生成的内容
    tab1, tab2 = st.tabs(["📄 完整计划", "🎯 核心亮点"])

    with tab1:
        st.markdown(PREGENERATED_RESULT)

        # 下载按钮
        st.download_button(
            label="⬇️ 下载完整计划（Markdown）",
            data=PREGENERATED_RESULT,
            file_name="李晓明_个性化30-60-90天计划.md",
            mime="text/markdown",
            use_container_width=True
        )

    with tab2:
        st.markdown("### 🔍 AI识别的核心洞察")

        col_insight1, col_insight2 = st.columns(2)

        with col_insight1:
            st.markdown("#### 🔵 优势识别")
            st.success("""
**Top 3 优势:**
1. **学科专业功底扎实** (92%)
   - 直接承担核心课程设计
   - 提高评估任务难度

2. **学习能力与主动性强** (88%)
   - 加速进入跨学科项目
   - 承担教研分享角色

3. **技术工具应用意识** (75%)
   - 引导探索数字化工具
   - 成为数字化教学先锋
""")

        with col_insight2:
            st.markdown("#### 🟡 短板防护")
            st.warning("""
**关键风险点:**
1. **课堂管理经验不足**
   - 前2周导师陪同上课
   - 提供结构化工具支持

2. **家校沟通待培养**
   - 提供3种场景话术模板
   - 每次沟通前预演
""")

        st.markdown("---")
        st.markdown("#### 💡 个性化调整示例")

        st.info("""
**标准要求:** 完成4个单元的教学设计（新教师通用）

**AI个性化调整:**
- ✅ **降低拆分粒度**: 由于专业功底强，可直接设计完整单元（不分拆小节）
- ✅ **增加创新空间**: 鼓励融入数字化工具（GeoGebra互动演示）
- 💡 **导师策略**: 前2周审核，第3-4周独立设计

**调整依据:**
- 胜任力匹配: "学科专业知识" 维度 92%
- 面试评价: "试讲逻辑清晰、概念准确"
- 简历亮点: "数学专业 + 教育学辅修"
""")

    st.markdown("---")

    # 操作按钮
    col_action1, col_action2, col_action3 = st.columns(3)

    with col_action1:
        if st.button("🔄 重新生成", use_container_width=True):
            st.session_state.step = 0
            st.session_state.show_result = False
            st.rerun()

    with col_action2:
        if st.button("📤 导出Word", use_container_width=True):
            st.info("💡 下载Markdown后可用Pandoc转换为Word格式")

    with col_action3:
        if st.button("📊 查看数据分析", use_container_width=True):
            st.info("🔜 高级分析功能开发中...")

# 页脚
st.markdown("---")
col_footer1, col_footer2, col_footer3 = st.columns([1, 2, 1])
with col_footer2:
    st.caption("🎯 HR个性化PD生成器 v0.2 | Demo版 | 黑客松专用")
    st.caption("💜 Powered by AI (Claude) | Made with ❤️ by ZWIE")
