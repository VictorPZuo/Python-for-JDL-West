# -*- coding: utf-8 -*-
"""
Streamlit App: 上架盘点表生成器（可选差异明细表 + 随机抽样）

功能：
- 上传 1) 前一日上架结果表（必选）
- 上传 2) 差异明细表（可选）：若未上传，则跳过“剔除差异储位”步骤，继续后续抽样
- 选择每个上架员抽样条数 k（可复现的随机种子 seed）
- 导出：仅下载【上架盘点表】（单 Sheet）

运行：
1) pip install streamlit pandas openpyxl numpy
2) streamlit run putaway_check_app_only_checktable_optional_diff.py
"""

import io
from typing import Tuple, Set, List

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="上架盘点表生成器", layout="wide")


# =========================
# 工具函数
# =========================
def _strip_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def read_excel_from_upload(uploaded_file, sheet_name=0) -> pd.DataFrame:
    return _strip_columns(pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl"))


def detect_diff_location_column(df_diff: pd.DataFrame) -> str:
    """
    自动识别差异明细表中“储位列”字段：
    优先：储位列
    其次：储位 / 储位编码 / 储位号 / 库位 / 库位编码
    否则：第一个包含“储位/库位”的列
    """
    cols = list(df_diff.columns)

    if "储位列" in cols:
        return "储位列"

    for cand in ["储位", "储位编码", "储位号", "库位", "库位编码"]:
        if cand in cols:
            return cand

    possible = [c for c in cols if ("储位" in c) or ("库位" in c)]
    if possible:
        return possible[0]

    raise ValueError(f"差异明细表中未找到包含“储位/库位”的列。现有列：{cols}")


def build_excluded_locations(df_diff: pd.DataFrame, loc_col: str) -> Set[str]:
    values = (
        df_diff[loc_col]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", np.nan)
        .dropna()
        .unique()
        .tolist()
    )
    return set(values)


def clean_source_data(
    df: pd.DataFrame,
    excluded_locations: Set[str],
    qty_limit: float = 50,
) -> pd.DataFrame:
    """
    步骤一：清洗源数据
    固定规则：
    1) 作业类型 == 采购进货
    2) 储区号 != R
    3) 上架量 <= 阈值（并可解析为数值）
    可选规则：
    4) 若 excluded_locations 非空：剔除储位编码在差异表中的记录
    """
    required = ["作业类型", "储区号", "上架量", "储位编码", "上架员"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"源文件缺少必要字段：{missing}")

    out = df.copy()
    out = out[out["作业类型"] == "采购进货"].copy()
    out = out[out["储区号"] != "R"].copy()

    out["上架量"] = pd.to_numeric(out["上架量"], errors="coerce")
    out = out[out["上架量"].notna() & (out["上架量"] <= qty_limit)].copy()

    out["储位编码"] = out["储位编码"].astype(str).str.strip()

    # 可选：剔除差异储位
    if excluded_locations:
        out = out[~out["储位编码"].isin(excluded_locations)].copy()

    return out.reset_index(drop=True)


def create_check_table(df_clean: pd.DataFrame) -> pd.DataFrame:
    """
    步骤二：生成上架盘点表（首列为上架员唯一值并剔除特定账号）
    """
    check = (
        df_clean[["上架员"]]
        .dropna()
        .drop_duplicates()
    )
    check = check[check["上架员"] != "xiao.han.1@jd.com"].copy()
    check = check.sort_values("上架员").reset_index(drop=True)
    return check


def sample_locations_to_check_table(
    df_clean: pd.DataFrame,
    check_table: pd.DataFrame,
    excluded_locations: Set[str],
    k: int,
    seed: int = 42,
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    步骤三：为每个上架员随机抽样 k 条储位编码，并写入储位编码_1..k
    规则：
    1) 抽样明细的上架员 != 当前上架员
    2) 已抽中的明细行不再参与后续抽样（全局不重复）
    3) 若 excluded_locations 非空：抽样明细的储位编码 不在差异明细表储位列中
    """
    rng = np.random.default_rng(seed)

    pool = df_clean.copy()
    pool["_row_id"] = pool.index

    result = check_table.copy()
    col_names = [f"储位编码_{i}" for i in range(1, k + 1)]
    for c in col_names:
        result[c] = pd.Series([None] * len(result), dtype="object")

    shortage: List[Tuple[str, int]] = []

    pool_loc_norm = pool["储位编码"].astype(str).str.strip()
    pool_user = pool["上架员"]

    for i, user in enumerate(result["上架员"].tolist()):
        eligible_mask = (pool_user != user)
        if excluded_locations:
            eligible_mask = eligible_mask & (~pool_loc_norm.isin(excluded_locations))

        eligible = pool.loc[eligible_mask]

        if eligible.empty:
            shortage.append((user, 0))
            continue

        n = min(k, len(eligible))
        sampled_ids = rng.choice(eligible["_row_id"].to_numpy(), size=n, replace=False)

        sampled_rows = eligible.set_index("_row_id").loc[sampled_ids]
        locs = sampled_rows["储位编码"].astype(str).str.strip().tolist()

        for j, val in enumerate(locs):
            result.at[i, col_names[j]] = val

        if n < k:
            shortage.append((user, n))

        # 全局去重：移除已抽取行
        remove_mask = pool["_row_id"].isin(sampled_ids)
        pool = pool.loc[~remove_mask].copy()
        pool_loc_norm = pool["储位编码"].astype(str).str.strip()
        pool_user = pool["上架员"]

    shortage_df = pd.DataFrame(shortage, columns=["上架员", "实际抽样条数"]).sort_values("实际抽样条数")
    return result, shortage_df


def add_inventory_content_column(result: pd.DataFrame, k: int) -> pd.DataFrame:
    """
    步骤四：新增“盘点内容列” = 所有储位编码列用英文逗号连接
    """
    out = result.copy()
    loc_cols = [f"储位编码_{i}" for i in range(1, k + 1)]
    out["盘点内容列"] = out[loc_cols].apply(
        lambda row: ",".join([str(x) for x in row if pd.notna(x) and str(x).strip() != ""]),
        axis=1
    )
    return out


def to_excel_bytes_single_sheet(df: pd.DataFrame, sheet_name: str = "上架盘点表") -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
    return output.getvalue()


# =========================
# UI
# =========================
st.title("上架盘点表生成器（Streamlit）")

with st.sidebar:
    st.header("上传文件")
    source_upload = st.file_uploader("1) 上传“前一日上架结果表”（必选）", type=["xlsx"])
    diff_upload = st.file_uploader("2) 上传“差异明细表”（可选）", type=["xlsx"])

    st.header("抽样参数")
    k = st.number_input("每个上架员抽样条数", min_value=1, max_value=50, value=10, step=1)
    seed = st.number_input("随机种子", min_value=0, max_value=10_000_000, value=42, step=1)

    st.header("清洗规则")
    qty_limit = st.number_input("上架量阈值（<=）", min_value=1.0, max_value=10_000.0, value=50.0, step=1.0)

run_btn = st.button("运行生成", type="primary", disabled=not source_upload)

if run_btn:
    try:
        df_source = read_excel_from_upload(source_upload)

        # 差异明细表：可选
        excluded_locations: Set[str] = set()
        if diff_upload is not None:
            df_diff = read_excel_from_upload(diff_upload)
            diff_loc_col = detect_diff_location_column(df_diff)
            excluded_locations = build_excluded_locations(df_diff, diff_loc_col)
            st.sidebar.success(f"已启用差异储位剔除（列：{diff_loc_col}，数量：{len(excluded_locations)}）")
        else:
            st.sidebar.info("未上传差异明细表：将跳过差异储位剔除。")

        df_clean = clean_source_data(df_source, excluded_locations, qty_limit=float(qty_limit))
        check_table = create_check_table(df_clean)

        sampled_table, shortage_df = sample_locations_to_check_table(
            df_clean, check_table, excluded_locations, k=int(k), seed=int(seed)
        )
        final_table = add_inventory_content_column(sampled_table, k=int(k))

        st.subheader("结果预览（前20行）")
        st.dataframe(final_table.head(20), use_container_width=True)

        if not shortage_df.empty:
            st.subheader("抽样不足名单（若有）")
            st.warning("存在上架员抽样不足（通常因为可抽样池不足）")
            st.dataframe(shortage_df, use_container_width=True)

        st.subheader("下载（仅上架盘点表）")
        export_bytes = to_excel_bytes_single_sheet(final_table)

        st.download_button(
            label="下载 上架盘点表.xlsx",
            data=export_bytes,
            file_name="上架盘点表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"运行失败：{e}")
else:
    st.info("请先上传“前一日上架结果表”，再点击运行。")
