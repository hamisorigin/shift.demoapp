import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="シフト入力テンプレート生成", layout="wide")
st.title("📅 シフト入力テンプレート自動生成アプリ")

st.markdown("""
このアプリでは、以下の情報を入力すると  
**「シフト自動作成汎用アプリ-入力表-1か月.xlsx」** と同じ形式のテンプレートExcelが生成されます。  
""")

# === 基本情報入力 ===
st.header("① 基本情報入力")

col1, col2, col3 = st.columns(3)
with col1:
    employees_text = st.text_area("従業員名（カンマ区切り）", "あ,い,う,え,お")
with col2:
    patterns_text = st.text_area("勤務パターン（カンマ区切り）", "早番,遅番")
with col3:
    attributes_text = st.text_area("属性（カンマ区切り）", "白,黒")

num_days = st.number_input("日数", min_value=1, max_value=31, value=30)

# === 入力データ整理 ===
I = [i.strip() for i in employees_text.split(",") if i.strip()]
T = [t.strip() for t in patterns_text.split(",") if t.strip()]
A = [a.strip() for a in attributes_text.split(",") if a.strip()]
D = [i+1 for i in range(num_days)]

# === 生成ボタン ===
if st.button("テンプレートExcelを生成"):
    st.success("✅ テンプレートを作成しました！")

    # 出勤可能日
    df_availability = pd.DataFrame(1, index=I, columns=D)
    df_availability.loc[:, :] = ""
    df_availability.index.name = "従業員"

    # 勤務可能パターン
    df_pattern = pd.DataFrame(1, index=I, columns=T)
    df_pattern.loc[:, :] = ""
    df_pattern.index.name = "従業員"

    # 勤務日数上下限
    df_limits = pd.DataFrame({"従業員": I, "下限": [0]*len(I), "上限": [num_days]*len(I)})

    # 従業員能力表
    df_ability = pd.DataFrame(1, index=I, columns=A)
    df_ability.loc[:, :] = ""
    df_ability.index.name = "従業員"

    # 属性ごとの必要点数
    df_need_attr = pd.DataFrame(1, index=D, columns=A)
    df_need_attr.loc[:, :] = ""
    df_need_attr.index.name = "日付"

    # 必要勤務人数
    df_need_pattern = pd.DataFrame(1, index=D, columns=T)
    df_need_pattern.loc[:, :] = ""
    df_need_pattern.index.name = "日付"

    # === Excel出力 ===
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_availability.to_excel(writer, sheet_name='出勤可能日')
        df_pattern.to_excel(writer, sheet_name='勤務可能パターン')
        df_limits.to_excel(writer, sheet_name='勤務日数上下限', index=False)
        df_ability.to_excel(writer, sheet_name='従業員能力表')
        df_need_attr.to_excel(writer, sheet_name='属性ごとの必要点数')
        df_need_pattern.to_excel(writer, sheet_name='必要勤務人数')

    st.download_button(
        label="📥 テンプレートExcelをダウンロード",
        data=output.getvalue(),
        file_name="シフト自動作成汎用アプリ-入力表-1か月.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.info("Excelをダウンロードして内容を入力後、最適化アプリにアップロードしてください。")
