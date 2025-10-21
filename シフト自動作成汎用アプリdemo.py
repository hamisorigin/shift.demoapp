import streamlit as st
import pandas as pd
from mip import *
import io
import tempfile
import os

# ==============================
# ページ設定
# ==============================
st.set_page_config(page_title="シフト自動作成アプリ", layout="wide")

page = st.sidebar.radio("ページを選択", ["テンプレート作成", "シフト最適化"])

# ==============================
# ページ① テンプレート作成
# ==============================
if page == "テンプレート作成":
    st.title("📋 シフト入力テンプレート自動生成")

    col1, col2, col3 = st.columns(3)
    with col1:
        employees_text = st.text_area("従業員名（カンマ区切り）", "あ,い,う,え,お")
    with col2:
        patterns_text = st.text_area("勤務パターン（カンマ区切り）", "早番,遅番")
    with col3:
        attributes_text = st.text_area("属性（カンマ区切り）", "白,黒")

    num_days = st.number_input("日数", min_value=1, max_value=31, value=30)

    I = [i.strip() for i in employees_text.split(",") if i.strip()]
    T = [t.strip() for t in patterns_text.split(",") if t.strip()]
    A = [a.strip() for a in attributes_text.split(",") if a.strip()]
    D = [i+1 for i in range(num_days)]

    if st.button("📄 テンプレートExcelを生成"):
        st.success("✅ テンプレートを作成しました！")

        df_availability = pd.DataFrame("", index=I, columns=D)
        df_availability.index.name = "従業員"

        df_pattern = pd.DataFrame("", index=I, columns=T)
        df_pattern.index.name = "従業員"

        df_limits = pd.DataFrame({"従業員": I, "下限": [0]*len(I), "上限": [num_days]*len(I)})

        df_ability = pd.DataFrame("", index=I, columns=A)
        df_ability.index.name = "従業員"

        df_need_attr = pd.DataFrame("", index=D, columns=A)
        df_need_attr.index.name = "日付"

        df_need_pattern = pd.DataFrame("", index=D, columns=T)
        df_need_pattern.index.name = "日付"

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

# ==============================
# ページ② シフト最適化
# ==============================
elif page == "シフト最適化":
    st.title("⚙️ シフト自動作成（MIP版）")

    uploaded_file = st.file_uploader("📤 Excelファイルをアップロード", type=["xlsx"])

    # ======= 入力検証 =======
    def validate_input_consistency(n, r, D, A, T):
        inconsistencies = []
        for d in D:
            total_required_staff = sum(r.get((d, t), 0) for t in T)
            total_required_points = sum(n.get((d, a), 0) for a in A)
            if total_required_points > total_required_staff:
                inconsistencies.append((d, total_required_staff, total_required_points))
        return inconsistencies

    # ======= 最適化処理 =======
    def run_shift_optimization(file_path):
        filename = file_path

        def extract_sheet_data(file, sheet_name):
            df = pd.read_excel(file, sheet_name=sheet_name, header=None)
            row_labels = df.iloc[1:, 0].dropna().tolist()
            col_raw = df.iloc[0, 1:].dropna().tolist()
            col_labels = [
                float(col) if isinstance(col, (int, float, pd.Int64Dtype().type)) and not isinstance(col, str) else col
                for col in col_raw
            ]
            result = []
            for i, row in enumerate(row_labels, start=1):
                for j, col in enumerate(col_labels, start=1):
                    value = df.iat[i, j]
                    if pd.notna(value):
                        result.append((row, col, float(value)))
            return result

        availability_list = extract_sheet_data(filename, '出勤可能日')
        pattern_list = extract_sheet_data(filename, '勤務可能パターン')
        limitdays_list = extract_sheet_data(filename, '勤務日数上下限')
        employeeability_list = extract_sheet_data(filename, '従業員能力表')
        needwork_list = extract_sheet_data(filename, '属性ごとの必要点数')
        required_staff_list = extract_sheet_data(filename, '必要勤務人数')

        I = sorted(set([row[0] for row in availability_list]))
        D = sorted(set([row[1] for row in availability_list]))
        T = sorted(set([row[1] for row in pattern_list]))
        A = sorted(set([row[1] for row in employeeability_list]))

        k = {(i, d): 0 for i in I for d in D}
        for i, d, val in availability_list:
            k[i, d] = int(val)

        g = {(i, t): 0 for i in I for t in T}
        for i, t, val in pattern_list:
            g[i, t] = int(val)

        l_min = {i: 0 for i in I}
        l_max = {i: len(D) for i in I}
        for i, _, val in limitdays_list:
            if val >= 10:
                l_max[i] = int(val)
            else:
                l_min[i] = int(val)

        s = {(i, a): 0 for i in I for a in A}
        for i, a, val in employeeability_list:
            s[i, a] = float(val)

        n = {(d, a): 0 for d in D for a in A}
        for d, a, val in needwork_list:
            n[d, a] = float(val)

        r = {(d, t): 0 for d in D for t in T}
        for d, t, val in required_staff_list:
            r[d, t] = int(val)

        inconsistencies = validate_input_consistency(n, r, D, A, T)
        if inconsistencies:
            msg = "⚠️ 以下の日付で「必要点数 > 必要人数」となっています：\n\n"
            for d, staff, points in inconsistencies:
                msg += f"・{d}日目: 必要人数={staff}, 必要点数={points}\n"
            st.warning(msg)

        # --- MIP モデル（略: あなたの既存コードそのまま） ---
        # ...（制約・目的関数・ソルバー呼び出し部は変更なし）

        # ダミーで返却（ここに最終DataFrameを入れる）
        return output, df_shift, df_attribute, df_pattern, df_total_workdays, df_dev

    # ======= UI操作 =======
    if uploaded_file:
        st.success("✅ ファイルを読み込みました！")

        if st.button("最適化を実行"):
            with st.spinner("最適化中...（数分かかる場合があります）"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(uploaded_file.read())
                    tmp_path = tmp.name

                result = run_shift_optimization(tmp_path)
                os.remove(tmp_path)

                if result:
                    output, df_shift, df_attr, df_pat, df_days, df_dev = result

                    st.success("✅ 最適化が完了しました！")

                    # ✅ Excelダウンロード
                    st.download_button(
                        "📥 結果Excelをダウンロード",
                        data=output.getvalue(),
                        file_name="シフト出力結果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    # ✅ 結果プレビュー
                    tab1, tab2, tab3, tab4, tab5 = st.tabs([
                        "📋 割り当て結果",
                        "📊 属性点数確認",
                        "👥 パターン人数確認",
                        "🗓 勤務日数集計",
                        "⚖ 属性偏り確認"
                    ])
                    with tab1:
                        st.dataframe(df_shift)
                    with tab2:
                        st.dataframe(df_attr)
                    with tab3:
                        st.dataframe(df_pat)
                    with tab4:
                        st.dataframe(df_days)
                    with tab5:
                        st.dataframe(df_dev)

                else:
                    st.error("❌ 解が見つかりませんでした。制約条件を確認してください。")
