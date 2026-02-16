# app.py
# -*- coding: utf-8 -*-
import os
import io
import pandas as pd
import streamlit as st
from datetime import date

st.set_page_config(page_title="UPPH by Location 报表生成器", layout="wide")
st.title("🗂️ UPPH by Location 报表生成器")
st.caption("上传 Excel 文件前先选择日期，系统将按 1–11 步自动处理，并以所选日期命名导出文件。")

# -------------------------
# 1) 选择日期（用于导出文件名）
# -------------------------
st.subheader("① 选择日期")
selected_date = st.date_input("请选择日期（用于导出文件名）", value=date.today())
# 文件名使用 MMDD（与“0911”示例一致）
date_str = selected_date.strftime("%m%d")  # 例如 2025-09-11 -> "0911"
st.write(f"将用于文件名：**UPPH by Location {date_str}.xlsx**")

# -------------------------
# 2) 上传文件
# -------------------------
st.subheader("② 上传 Excel 文件")
uploaded = st.file_uploader("📂 请选择要处理的 Excel 文件（.xlsx）", type=["xlsx"])

# -------------------------
# 工具函数
# -------------------------
def load_excel_from_bytes(file_bytes: bytes) -> pd.DataFrame:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    return pd.read_excel(xls, sheet_name=xls.sheet_names[0])

def process(df_raw: pd.DataFrame) -> pd.DataFrame:
    # Step 2: 删除“储位”=R01011
    df = df_raw[df_raw["储位"] != "R01011"].copy()

    # Step 3: 新增“UPPH by Location”= 任务单号 + 储位
    df["UPPH by Location"] = df["任务单号"].astype(str) + df["储位"].astype(str)

    # Step 4: 提取“拣货完成时间”的小时
    df["拣货完成时间"] = pd.to_datetime(df["拣货完成时间"], errors="coerce")
    df["拣货完成时间_Hour"] = df["拣货完成时间"].dt.hour

    # Step 5: 透视表（行=邮箱，列=小时，值=UPPH by Location 的非重复计数）
    pivot = pd.pivot_table(
        df,
        index="邮箱",
        columns="拣货完成时间_Hour",
        values="UPPH by Location",
        aggfunc=pd.Series.nunique,
        fill_value=0
    ).reset_index()

    # Step 6: 新增“姓名”列（邮箱 -> 姓名）
    email_to_name = (
        df.dropna(subset=["邮箱", "姓名"])
          .drop_duplicates("邮箱")
          .set_index("邮箱")["姓名"]
          .to_dict()
    )
    pivot.insert(1, "姓名", pivot["邮箱"].map(email_to_name))

    # Step 7: 新增“触碰储位总数”
    hour_cols = [c for c in pivot.columns if isinstance(c, (int, float))]
    pivot["触碰储位总数"] = pivot[hour_cols].sum(axis=1)

    # Step 8: 新增“工作时长”（非零小时段数）
    pivot["工作时长"] = (pivot[hour_cols] > 0).sum(axis=1)

    # Step 9: 新增“UPPH by Location (Avg)”= 触碰储位总数 / 工作时长（两位小数）
    pivot["UPPH by Location (Avg)"] = (
        pivot["触碰储位总数"] / pivot["工作时长"].replace(0, pd.NA)
    ).round(2)

    # Step 10: 按平均值降序排序
    pivot_sorted = pivot.sort_values(
        by="UPPH by Location (Avg)", ascending=False, na_position="last"
    ).reset_index(drop=True)

    return pivot_sorted

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Result")
    buf.seek(0)
    return buf.read()

# -------------------------
# 主逻辑
# -------------------------
if uploaded is not None:
    try:
        # 读取
        df_raw = load_excel_from_bytes(uploaded.read())
        st.success("✅ 文件已上传并读取成功")
        st.markdown("### 源数据预览")
        st.dataframe(df_raw.head(20), use_container_width=True)

        # 处理
        with st.spinner("正在处理数据..."):
            result = process(df_raw)

        st.markdown("### 最终结果（已按 UPPH by Location (Avg) 降序）")
        st.dataframe(result, use_container_width=True)

        # 导出：下载 & 保存到桌面（文件名包含所选日期）
        st.markdown("---")
        file_name = f"UPPH by Location {date_str}.xlsx"

        # 下载
        excel_bytes = to_excel_bytes(result)
        st.download_button(
            label=f"⬇️ 下载结果：{file_name}",
            data=excel_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        # 保存到桌面
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        if st.button(f"💾 保存到桌面（{file_name}）"):
            try:
                output_path = os.path.join(desktop, file_name)
                result.to_excel(output_path, index=False)
                st.success(f"✅ 已保存到桌面：{output_path}")
            except Exception as e:
                st.error(f"❌ 保存失败：{e}")

    except Exception as e:
        st.error(f"❌ 处理失败：{e}")
else:
    st.info("👆 请先选择上方日期，然后上传 Excel 文件。")