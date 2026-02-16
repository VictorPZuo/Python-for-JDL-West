import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
from typing import Optional


def classify_sku_fixed_window(
    df: pd.DataFrame,
    high_qty_threshold: int = 20,      # 高销量阈值（件）
    high_days_threshold: int = 7,      # 高销量天数阈值（天）
    mid_qty_threshold: int = 3,        # 中销量阈值（件）
    mid_days_threshold: int = 7,       # 中销量天数阈值（天）
    active_days_threshold: int = 5,    # 动销天数阈值（天）
    window_days: Optional[int] = 14,   # 统计最近 N 天，默认 14 天
) -> pd.DataFrame:
    """
    按 SKU 统计动销天数、高销量天数、中销量天数，并打 A/B/C/D 标签。

    必需列：商品编码、任务下发时间、预期拣货量
    可以选择只统计最近 window_days 天的数据（基于日期最大值向前回溯）
    """

    # 处理日期
    df['任务下发时间'] = pd.to_datetime(df['任务下发时间'])
    df['日期'] = df['任务下发时间'].dt.date

    # 如果指定了统计窗口天数，则只保留最近 window_days 天
    if window_days is not None:
        max_date = df['日期'].max()
        if pd.isna(max_date):
            # 没有有效日期，直接返回空
            return pd.DataFrame(columns=[
                'SKU', '动销天数',
                f'高销量天数(≥{high_qty_threshold}件)',
                f'中销量天数(≥{mid_qty_threshold}件)',
                '标签'
            ])
        cutoff_date = max_date - datetime.timedelta(days=window_days - 1)
        df = df[df['日期'] >= cutoff_date]

    # 如果过滤后没有数据，直接返回空表
    if df.empty:
        return pd.DataFrame(columns=[
            'SKU', '动销天数',
            f'高销量天数(≥{high_qty_threshold}件)',
            f'中销量天数(≥{mid_qty_threshold}件)',
            '标签'
        ])

    # 按 SKU + 日期 聚合每日拣货量
    df_daily = df.groupby(['商品编码', '日期'])['预期拣货量'].sum().reset_index()

    results = []

    # 每个 SKU 进行分档
    for sku, sku_df in df_daily.groupby('商品编码'):
        window = sku_df
        # 高销量天数：每日拣货量 ≥ high_qty_threshold
        days_ge_high = (window['预期拣货量'] >= high_qty_threshold).sum()
        # 中销量天数：每日拣货量 ≥ mid_qty_threshold
        days_ge_mid = (window['预期拣货量'] >= mid_qty_threshold).sum()
        # 动销天数：每日拣货量 > 0
        active_days = (window['预期拣货量'] > 0).sum()

        # 分档逻辑
        if days_ge_high >= high_days_threshold:
            label = 'A'
        elif days_ge_mid >= mid_days_threshold:
            label = 'B'
        elif active_days >= active_days_threshold:
            label = 'C'
        else:
            label = 'D'

        results.append({
            'SKU': sku,
            '动销天数': active_days,
            f'高销量天数(≥{high_qty_threshold}件)': days_ge_high,
            f'中销量天数(≥{mid_qty_threshold}件)': days_ge_mid,
            '标签': label
        })

    return pd.DataFrame(results)


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "SKU分档结果") -> bytes:
    """把 DataFrame 转成 Excel 二进制，方便下载"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output.getvalue()


def main():
    st.title("SKU 分档小程序（2周拣货数据 · V2.2）")

    st.markdown("""
    **使用说明：**
    1. 上传最近 2 周的拣货数据文件（Excel）  
       - 必须包含列：`商品编码`、`任务下发时间`、`预期拣货量`  
    2. 上传当前「库存表」（Excel）  
       - 必须包含列：`京东商品编码`、`货型`、`货主名称`（用于匹配）  
    3. 在左侧栏可以调整分档参数与统计窗口（默认最近 14 天）  
    4. 点击 **运行分档** 按钮  
    5. 程序会生成 **SKU分档结果表**，并执行步骤一：  
       - `SKU分档结果表.SKU` 关联 `库存表.京东商品编码`  
       - 把 `货型`、`货主名称` 两列添加到 `SKU分档结果表`
    """)

    # ========== 侧边栏：参数设置 ==========
    st.sidebar.header("分档参数设置")

    window_days = st.sidebar.number_input(
        "统计窗口（最近 N 天）",
        min_value=1,
        max_value=365,
        value=14,
        step=1,
        help="以数据中的最大日期为基准，向前回溯 N 天进行统计"
    )

    st.sidebar.subheader("高销量档（A）规则")
    high_qty_threshold = st.sidebar.number_input(
        "高销量阈值（件）",
        min_value=1,
        value=20,
        step=1
    )
    high_days_threshold = st.sidebar.number_input(
        "高销量天数阈值（天）",
        min_value=1,
        value=7,
        step=1
    )

    st.sidebar.subheader("中销量档（B）规则")
    mid_qty_threshold = st.sidebar.number_input(
        "中销量阈值（件）",
        min_value=1,
        value=3,
        step=1
    )
    mid_days_threshold = st.sidebar.number_input(
        "中销量天数阈值（天）",
        min_value=1,
        value=7,
        step=1
    )

    st.sidebar.subheader("低销量档（C）规则")
    active_days_threshold = st.sidebar.number_input(
        "动销天数阈值（天）",
        min_value=1,
        value=5,
        step=1
    )

    # ========== 首页：文件上传 ==========
    st.subheader("数据上传")

    col_up1, col_up2 = st.columns(2)

    with col_up1:
        picking_file = st.file_uploader(
            "上传「2周拣货数据」Excel 文件",
            type=["xlsx", "xls"],
            key="picking_uploader"
        )

    with col_up2:
        inventory_file = st.file_uploader(
            "上传「库存表」Excel 文件",
            type=["xlsx", "xls"],
            key="inventory_uploader"
        )

    # 运行按钮
    run = st.button("运行分档")

    if run:
        # 1) 检查拣货数据
        if picking_file is None:
            st.error("请先上传「2周拣货数据」文件，再点击运行。")
            return

        try:
            df_picking = pd.read_excel(picking_file)
        except Exception as e:
            st.error(f"读取「2周拣货数据」Excel 文件失败，请检查文件格式。错误信息：{e}")
            return

        required_cols = ["商品编码", "任务下发时间", "预期拣货量"]
        missing = [c for c in required_cols if c not in df_picking.columns]

        if missing:
            st.error(f"拣货数据中缺少必要列：{missing}。请检查源数据。")
            return

        # 2) 读取库存表（用于步骤一匹配）
        df_inventory = None
        if inventory_file is not None:
            try:
                df_inventory = pd.read_excel(inventory_file)
            except Exception as e:
                st.warning(f"读取「库存表」失败：{e}。本次将仅基于拣货数据做分档。")
        else:
            st.warning("未上传「库存表」，将无法执行步骤一匹配（货型、货主名称）。")

        # 3) 分档计算
        with st.spinner("正在进行 SKU 分档计算，请稍候…"):
            result = classify_sku_fixed_window(
                df_picking,
                high_qty_threshold=high_qty_threshold,
                high_days_threshold=high_days_threshold,
                mid_qty_threshold=mid_qty_threshold,
                mid_days_threshold=mid_days_threshold,
                active_days_threshold=active_days_threshold,
                window_days=window_days,
            )

        if result.empty:
            st.warning("在当前统计窗口内没有有效拣货数据，请检查日期范围或源数据内容。")
            return

        st.success("SKU 分档完成！（SKU分档结果表）")

        # ========== 步骤一：与库存表关联，加入“货型”“货主名称” ==========
        if df_inventory is not None:
            # 检查关键列是否存在
            inv_required_cols = ["京东商品编码", "货型", "货主名称"]
            inv_missing = [c for c in inv_required_cols if c not in df_inventory.columns]

            if inv_missing:
                st.warning(
                    f"库存表中缺少以下列：{inv_missing}，无法执行步骤一匹配（货型、货主名称）。"
                )
            else:
                # 复制一份，避免修改原 DataFrame 的列名
                df_inv_for_merge = df_inventory.copy()
                df_inv_for_merge = df_inv_for_merge.rename(columns={
                    "京东商品编码": "SKU"
                })

                # 只保留相关列并去重
                inv_subset = df_inv_for_merge[["SKU", "货型", "货主名称"]].drop_duplicates()

                # 左连接，确保 SKU 分档结果完整
                result = result.merge(inv_subset, on="SKU", how="left")

                st.info("已完成步骤一：成功将【货型】【货主名称】添加到 SKU 分档结果表。")

                # 库存表预览（可选）
                st.subheader("库存表预览（前 20 行）")
                st.dataframe(df_inventory.head(20), use_container_width=True)
        else:
            st.info("由于未上传库存表，本次未执行步骤一匹配（货型、货主名称）。")

        # ========== 分档概览与标签分布 ==========
        st.subheader("分档概览")

        total_skus = len(result)
        label_counts = result['标签'].value_counts().reindex(['A', 'B', 'C', 'D'], fill_value=0)
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("总 SKU 数", total_skus)
        col2.metric("A 档 SKU 数", int(label_counts.get('A', 0)))
        col3.metric("B 档 SKU 数", int(label_counts.get('B', 0)))
        col4.metric("C 档 SKU 数", int(label_counts.get('C', 0)))
        col5.metric("D 档 SKU 数", int(label_counts.get('D', 0)))

        st.markdown("**标签分布图：**")
        st.bar_chart(label_counts)

        # ========== 标签筛选 & 结果预览 ==========
        st.subheader("分档结果明细（含步骤一匹配结果）")

        available_labels = sorted(result['标签'].unique())
        selected_labels = st.multiselect(
            "按标签筛选（默认全部）",
            options=available_labels,
            default=available_labels
        )

        if selected_labels:
            filtered_result = result[result['标签'].isin(selected_labels)]
        else:
            filtered_result = result.copy()

        st.dataframe(filtered_result, use_container_width=True)

        # ========== 下载结果 ==========
        st.markdown("---")
        st.markdown("### 下载结果")

        excel_bytes = to_excel_bytes(result, sheet_name="SKU分档结果")
        st.download_button(
            label="下载完整分档结果（Excel，含货型 & 货主名称）",
            data=excel_bytes,
            file_name="SKU分档结果_含货型货主.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    main()
