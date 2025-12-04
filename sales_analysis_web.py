import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os

# 设置中文支持
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# Web页面标题
st.title("📊 销售数据分析工具")
st.subheader("2024年订单数据可视化分析")

# 1. 上传Excel文件
uploaded_file = st.file_uploader("选择订单Excel文件", type=["xlsx", "xls"])
if uploaded_file is not None:
    # 读取数据
    df = pd.read_excel(uploaded_file)
    st.success(f"✅ 成功读取文件：{uploaded_file.name}（总订单数：{len(df)}笔）")
    
    # 2. 数据清洗
    df_clean = df[df["订单金额"] > 0].copy()
    异常数 = len(df) - len(df_clean)
    if 异常数 > 0:
        st.warning(f"⚠️  过滤异常订单（金额≤0）：{异常数}笔，有效订单数：{len(df_clean)}笔")
    else:
        st.info(f"✅ 无异常订单，有效订单数：{len(df_clean)}笔")
    
    # 3. 核心指标计算
    总销售额 = df_clean["订单金额"].sum()
    平均订单金额 = df_clean["订单金额"].mean()
    高价值订单 = df_clean[df_clean["订单金额"] > 5000]
    
    # 显示核心指标（卡片式布局）
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("💰 总销售额", f"{总销售额:.2f}元")
    with col2:
        st.metric("📈 平均订单金额", f"{平均订单金额:.2f}元")
    with col3:
        st.metric("🏆 高价值订单数", f"{len(高价值订单)}笔")
    
    # 4. 图表展示（可切换）
    st.subheader("📊 数据可视化")
    图表类型 = st.selectbox("选择图表", ["客户销售额TOP10", "订单金额分布", "高价值订单占比"])
    
    if 图表类型 == "客户销售额TOP10":
        客户销售额 = df_clean.groupby("客户名称")["订单金额"].sum().sort_values(ascending=False).head(10)
        fig, ax = plt.subplots(figsize=(10, 5))
        客户销售额.plot(kind="bar", ax=ax, color="#3498DB")
        ax.set_title("客户销售额TOP10")
        ax.set_xlabel("客户名称")
        ax.set_ylabel("销售额（元）")
        ax.tick_params(axis='x', rotation=45)
        st.pyplot(fig)
    
    elif 图表类型 == "订单金额分布":
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.hist(df_clean["订单金额"], bins=20, color="#E74C3C", alpha=0.7)
        ax.axvline(平均订单金额, color="red", linestyle="--", label=f"平均金额：{平均订单金额:.0f}元")
        ax.set_title("订单金额分布")
        ax.set_xlabel("订单金额（元）")
        ax.set_ylabel("订单数量")
        ax.legend()
        st.pyplot(fig)
    
    elif 图表类型 == "高价值订单占比":
        分层销售额 = [高价值订单["订单金额"].sum(), 总销售额 - 高价值订单["订单金额"].sum()]
        分层标签 = ["高价值订单", "普通订单"]
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.pie(分层销售额, labels=分层标签, autopct="%1.1f%%", colors=["#E74C3C", "#27AE60"])
        ax.set_title("高价值订单销售额占比")
        ax.axis("equal")
        st.pyplot(fig)
    
    # 5. 客户分层分析
    st.subheader("👥 客户分层分析")
    客户总消费 = df_clean.groupby("客户名称")["订单金额"].sum()
    高消费客户 = 客户总消费[客户总消费 > 10000]
    中消费客户 = 客户总消费[(客户总消费 >= 3000) & (客户总消费 <= 10000)]
    低消费客户 = 客户总消费[客户总消费 < 3000]
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.metric("💎 高消费客户（>1万）", f"{len(高消费客户)}人")
        st.text(f"贡献销售额：{高消费客户.sum():.2f}元")
    with col_b:
        st.metric("🌟 中消费客户（3千-1万）", f"{len(中消费客户)}人")
        st.text(f"贡献销售额：{中消费客户.sum():.2f}元")
    with col_c:
        st.metric("🌱 低消费客户（<3千）", f"{len(低消费客户)}人")
        st.text(f"贡献销售额：{低消费客户.sum():.2f}元")
    
    # 6. 下载报告
    st.subheader("📥 下载分析报告")
    with pd.ExcelWriter("销售分析报告.xlsx", engine="openpyxl") as writer:
        核心指标 = pd.DataFrame({
            "指标名称": ["总订单数", "有效订单数", "总销售额", "平均订单金额", "高价值订单数"],
            "数值": [len(df), len(df_clean), 总销售额, 平均订单金额, len(高价值订单)]
        })
        核心指标.to_excel(writer, sheet_name="核心指标", index=False)
        客户销售额.reset_index().to_excel(writer, sheet_name="客户销售排行", index=False)
    
    with open("销售分析报告.xlsx", "rb") as f:
        st.download_button(
            label="下载Excel报告",
            data=f,
            file_name=f"销售分析报告_{uploaded_file.name.split('.')[0]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("请上传Excel订单文件开始分析（支持.xlsx/.xls格式）")