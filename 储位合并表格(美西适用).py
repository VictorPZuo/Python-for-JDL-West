import streamlit as st
import pandas as pd
from io import BytesIO


st.title("可合并储位筛选工具")

st.markdown("### 📂 上传数据")

# 上传 储位表
storage_file = st.file_uploader(
    "请上传【储位表】（Excel：.xlsx / .xls）",
    type=["xlsx", "xls"],
    key="storage_uploader"
)

# 上传 库存表
inventory_file = st.file_uploader(
    "请上传【库存表】（Excel：.xlsx / .xls 或 CSV）",
    type=["xlsx", "xls", "csv"],
    key="inventory_uploader"
)

# 选择仓号（筛选条件）
warehouse_option = st.selectbox(
    "请选择仓号（筛选条件）",
    ["LAX1", "LAX2", "LAX4", "LAX5"]
)

# 运行按钮
run_button = st.button("运行")


def read_inventory_file(file):
    """根据扩展名读取库存表"""
    name = file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    else:
        return pd.read_excel(file)


if run_button:
    if storage_file is None or inventory_file is None:
        st.error("⚠️ 请先上传【储位表】和【库存表】后再点击运行。")
    else:
        # =========================
        # 1. 读取原始数据
        # =========================
        storage_df = pd.read_excel(storage_file)
        inventory_df = read_inventory_file(inventory_file)

        # 必要字段校验
        must_inv_cols = [
            "京东商品编码", "货主名称", "储位编码",
            "库存量", "可用量", "货型",
            "长", "宽", "高"          # 商品尺寸（英寸）
        ]
        miss_inv = [c for c in must_inv_cols if c not in inventory_df.columns]
        if miss_inv:
            st.error(f"库存表缺少必要列：{miss_inv}，请检查后重新上传。")
            st.stop()

        must_sto_cols = [
            "储位编码", "储位规格", "层",
            "长", "宽", "高"          # 储位尺寸（毫米）
        ]
        miss_sto = [c for c in must_sto_cols if c not in storage_df.columns]
        if miss_sto:
            st.error(f"储位表缺少必要列：{miss_sto}，请检查后重新上传。")
            st.stop()

        # 只保留需要的列
        inventory_df = inventory_df[must_inv_cols].copy()

        # 数值列转为数值类型
        for col in ["库存量", "可用量", "长", "宽", "高"]:
            inventory_df[col] = pd.to_numeric(inventory_df[col], errors="coerce").fillna(0)

        for col in ["长", "宽", "高"]:
            storage_df[col] = pd.to_numeric(storage_df[col], errors="coerce").fillna(0)

        # =========================
        # 2. 仓号筛选规则 → L2-L4库位表
        # =========================
        def rule_LAX1(df: pd.DataFrame) -> pd.DataFrame:
            # 暂无规则 → 不过滤
            return df.copy()

        def rule_LAX2(df: pd.DataFrame) -> pd.DataFrame:
            # 储位规格 ∈ {CW05, CW06, CW08} 且 层 ∈ {2,3,4}
            return df[
                df["储位规格"].isin(["CW05", "CW06", "CW08"]) &
                df["层"].isin([2, 3, 4])
            ].copy()

        def rule_LAX4(df: pd.DataFrame) -> pd.DataFrame:
            return df.copy()

        def rule_LAX5(df: pd.DataFrame) -> pd.DataFrame:
            # 储位规格 ∈ {DCS00000001} 且 层 ∈ {2,3,4,5}
            return df[
                df["储位规格"].isin(["DCS00000001"]) &
                df["层"].isin([2, 3, 4, 5])
            ].copy()

        if warehouse_option == "LAX1":
            L2_L4库位表 = rule_LAX1(storage_df)
        elif warehouse_option == "LAX2":
            L2_L4库位表 = rule_LAX2(storage_df)
        elif warehouse_option == "LAX4":
            L2_L4库位表 = rule_LAX4(storage_df)
        else:  # LAX5
            L2_L4库位表 = rule_LAX5(storage_df)

        # =========================
        # 3. 按 L2-L4 库位过滤库存表  → 库存表_过滤后
        # =========================
        valid_locations = L2_L4库位表["储位编码"].unique()
        库存表_过滤后 = inventory_df[inventory_df["储位编码"].isin(valid_locations)].copy()

        # =========================
        # 4. 生成 SKU_众数表（只保留储位数>2 的 SKU）
        # =========================
        sku_counts = (
            库存表_过滤后.groupby("京东商品编码")["储位编码"]
            .count()
            .reset_index(name="储位数")
        )
        skus_gt2 = sku_counts[sku_counts["储位数"] > 2]["京东商品编码"]

        库存表_SKU大于2 = 库存表_过滤后[
            库存表_过滤后["京东商品编码"].isin(skus_gt2)
        ].copy()

        def get_mode(series: pd.Series):
            modes = series.mode()
            return modes.iloc[0] if not modes.empty else None

        SKU_众数表 = (
            库存表_SKU大于2.groupby("京东商品编码")["可用量"]
            .apply(get_mode)
            .reset_index(name="可用量_众数")
        )

        # =========================
        # 5. 生成可合并储位表（可用量 < 可用量_众数）
        # =========================
        库存表_带众数 = 库存表_过滤后.merge(
            SKU_众数表,
            on="京东商品编码",
            how="inner"
        )

        mask = 库存表_带众数["可用量"] < 库存表_带众数["可用量_众数"]
        可合并储位表 = 库存表_带众数[mask].copy()

        # =========================
        # 6. 计算储位利用率（修正版）
        #   6.1 使用【库存表】计算库存体积（英寸 → 立方米）
        #   6.2 使用【L2_L4库位表】计算储位体积（毫米 → 立方米）
        # =========================

        # 6.1 商品体积：库存表 长/宽/高 为英寸 → in³ → m³
        INCH3_TO_M3 = 0.0254 ** 3

        inv_vol = inventory_df.copy()
        inv_vol["单件体积_m3"] = (
            inv_vol["长"] * inv_vol["宽"] * inv_vol["高"] * INCH3_TO_M3
        )
        inv_vol["库存体积_m3"] = inv_vol["单件体积_m3"] * inv_vol["库存量"]

        储位_库存体积表 = (
            inv_vol.groupby("储位编码")["库存体积_m3"]
            .sum()
            .reset_index()
        )

        # 6.2 储位体积：储位表 长/宽/高 为毫米 → m → m³
        slot_vol = L2_L4库位表.copy()
        slot_vol["储位体积_m3"] = (
            (slot_vol["长"] / 1000.0) *
            (slot_vol["宽"] / 1000.0) *
            (slot_vol["高"] / 1000.0)
        )
        储位体积简表 = slot_vol[["储位编码", "储位体积_m3"]].copy()

        # 6.3 合并并计算利用率
        利用率表 = 储位_库存体积表.merge(
            储位体积简表,
            on="储位编码",
            how="left"
        )

        denom = 利用率表["储位体积_m3"].replace(0, pd.NA)
        利用率表["储位利用率"] = 利用率表["库存体积_m3"] / denom
        利用率表["储位利用率"] = 利用率表["储位利用率"].fillna(0)

        # 换算为百分数（0–100，保留 2 位小数）
        利用率表["储位利用率"] = (利用率表["储位利用率"] * 100).round(2)

        # 回填至可合并储位表
        可合并储位表 = 可合并储位表.merge(
            利用率表[["储位编码", "储位利用率"]],
            on="储位编码",
            how="left"
        )

        # =========================
        # 7. 按储位统计 京东商品编码数量，并回填
        #    （这里建议用 库存表_过滤后，保证只统计当前仓号+L2-L4 范围的 SKU 数）
        # =========================
        slot_sku_cnt = (
            库存表_过滤后.groupby("储位编码")["京东商品编码"]
            .nunique()
            .reset_index(name="京东商品编码数量")
        )

        可合并储位表 = 可合并储位表.merge(
            slot_sku_cnt,
            on="储位编码",
            how="left"
        )

        # =========================
        # 8. 排序：
        #    1) 每个京东商品编码在表中的条数（多→少）
        #    2) 储位利用率（少→多）
        #    3) 京东商品编码数量（少→多）
        # =========================
        sort_df = 可合并储位表.copy()

        sort_df["SKU条数"] = sort_df.groupby("京东商品编码")["储位编码"].transform("count")

        sort_df = sort_df.sort_values(
            by=["SKU条数", "储位利用率", "京东商品编码数量"],
            ascending=[False, True, True]
        ).drop(columns=["SKU条数"])

        # 最终列顺序
        final_cols = [
            "京东商品编码",
            "货主名称",
            "储位编码",
            "库存量",
            "可用量",
            "货型",
            "储位利用率",
            "京东商品编码数量",
        ]
        sort_df = sort_df[final_cols]

        # =========================
        # 9. 仅展示最终【可合并储位表】
        # =========================
        st.subheader("最终【可合并储位表】")
        st.write(f"仓号：{warehouse_option}，共 {sort_df.shape[0]} 条记录")
        st.dataframe(sort_df)

        # =========================
        # 10. 导出 Excel（下载按钮）
        # =========================
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            sort_df.to_excel(writer, index=False, sheet_name="可合并储位表")
        output.seek(0)

        st.download_button(
            label="📥 下载 可合并储位表.xlsx",
            data=output,
            file_name=f"可合并储位表_{warehouse_option}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
