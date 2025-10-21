import streamlit as st
import pandas as pd
import pulp
import io
import tempfile
import os

# ==============================
# ページ設定
# ==============================
st.set_page_config(page_title="シフト自動作成アプリ", layout="wide")

# --- ページ選択 ---
page = st.sidebar.radio("ページを選択", ["テンプレート作成", "シフト最適化"])

# ==============================
# ページ① テンプレート作成
# ==============================
if page == "テンプレート作成":
    st.title("📋 シフト入力テンプレート自動生成")

    st.markdown("""
    **このページでは、入力Excelテンプレートを自動生成します。**  
    生成後、Excelに必要情報を入力して「シフト最適化」ページでアップロードしてください。
    """)

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

        # 出勤可能日
        df_availability = pd.DataFrame("", index=I, columns=D)
        df_availability.index.name = "従業員"

        # 勤務可能パターン
        df_pattern = pd.DataFrame("", index=I, columns=T)
        df_pattern.index.name = "従業員"

        # 勤務日数上下限
        df_limits = pd.DataFrame({"従業員": I, "下限": [0]*len(I), "上限": [num_days]*len(I)})

        # 従業員能力表
        df_ability = pd.DataFrame("", index=I, columns=A)
        df_ability.index.name = "従業員"

        # 属性ごとの必要点数
        df_need_attr = pd.DataFrame("", index=D, columns=A)
        df_need_attr.index.name = "日付"

        # 必要勤務人数
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

        st.info("💡 Excelをダウンロードし、必要情報を入力後「シフト最適化」ページでアップロードしてください。")

# ==============================
# ページ② シフト最適化
# ==============================
elif page == "シフト最適化":
    st.title("⚙️ シフト自動作成（PuLP版）")

    st.markdown("""
    **以下の手順でシフトを最適化します：**  
    1. 入力表（テンプレートに記入したExcel）をアップロード  
    2. 「最適化を実行」をクリック  
    3. 結果Excelをダウンロードできます
    """)

    uploaded_file = st.file_uploader("📤 Excelファイルをアップロード", type=["xlsx"])

    # ======= 入力検証関数 =======
    def validate_input_consistency(n, r, D, A, T):
        """必要人数と必要点数の整合性チェック"""
        inconsistencies = []
        for d in D:
            total_required_staff = sum(r.get((d, t), 0) for t in T)
            total_required_points = sum(n.get((d, a), 0) for a in A)
            if total_required_points > total_required_staff:
                inconsistencies.append((d, total_required_staff, total_required_points))
        return inconsistencies

    # ======= 最適化処理 (PuLP) =======
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

        # --- データ読み込み ---
        availability_list = extract_sheet_data(filename, '出勤可能日')
        pattern_list = extract_sheet_data(filename, '勤務可能パターン')
        limitdays_list = extract_sheet_data(filename, '勤務日数上下限')
        employeeability_list = extract_sheet_data(filename, '従業員能力表')
        needwork_list = extract_sheet_data(filename, '属性ごとの必要点数')
        required_staff_list = extract_sheet_data(filename, '必要勤務人数')

        # --- 集合 ---
        I = sorted(set([row[0] for row in availability_list]))
        D = sorted(set([row[1] for row in availability_list]))
        T = sorted(set([row[1] for row in pattern_list]))
        A = sorted(set([row[1] for row in employeeability_list]))

        # --- 定数 ---
        k = {(i, d): 0 for i in I for d in D}
        for i, d, val in availability_list:
            if d in D:
                k[i, d] = int(val)

        g = {(i, t): 0 for i in I for t in T}
        for i, t, val in pattern_list:
            if t in T:
                g[i, t] = int(val)

        l_min = {i: 0 for i in I}
        l_max = {i: len(D) for i in I}
        for i, _, val in limitdays_list:
            if i in I:
                if val >= 10:
                    l_max[i] = int(val)
                else:
                    l_min[i] = int(val)

        s = {(i, a): 0 for i in I for a in A}
        for i, a, val in employeeability_list:
            if a in A:
                s[i, a] = float(val)

        n = {(d, a): 0 for d in D for a in A}
        for d, a, val in needwork_list:
            if d in D and a in A:
                n[d, a] = float(val)

        r = {(d, t): 0 for d in D for t in T}
        for d, t, val in required_staff_list:
            if d in D and t in T:
                r[d, t] = int(val)

        # --- 入力整合性チェック ---
        inconsistencies = validate_input_consistency(n, r, D, A, T)
        if inconsistencies:
            msg = "⚠️ 以下の日付で「必要点数 > 必要人数」となっています。\n最適化を実行するとペナルティが発生します。\n\n"
            for d, staff, points in inconsistencies:
                msg += f"・{d}日目: 必要人数={staff}, 必要点数={points}\n"
            st.warning(msg)

        # --- PuLP モデル定義 ---
        prob = pulp.LpProblem("ShiftScheduling_4D_PuLP", pulp.LpMinimize)

        # 決定変数 x[i,d,t,a] ∈ {0,1}
        x = {}
        for i in I:
            for d in D:
                for t in T:
                    for a in A:
                        name = f"x_{str(i)}_{str(d)}_{str(t)}_{str(a)}"
                        x[(i, d, t, a)] = pulp.LpVariable(name, cat="Binary")

        # ペナルティ変数
        shortfall_attr = {}
        shortfall_pat = {}
        for d in D:
            for a in A:
                shortfall_attr[(d, a)] = pulp.LpVariable(f"shortfall_attr_{d}_{a}", lowBound=0, cat="Continuous")
        for d in D:
            for t in T:
                shortfall_pat[(d, t)] = pulp.LpVariable(f"shortfall_pat_{d}_{t}", lowBound=0, cat="Continuous")

        # 制約
        for i in I:
            for d in D:
                for t in T:
                    for a in A:
                        prob += x[(i, d, t, a)] <= k.get((i, d), 0)
                        prob += x[(i, d, t, a)] <= g.get((i, t), 0)

        for i in I:
            vars_i = [x[(i, d, t, a)] for d in D for t in T for a in A]
            prob += pulp.lpSum(vars_i) >= l_min.get(i, 0)
            prob += pulp.lpSum(vars_i) <= l_max.get(i, len(D))

        try:
            D_numeric = sorted([int(d) for d in D])
            for i in I:
                for idx in range(len(D_numeric) - 4):
                    window_days = D_numeric[idx:idx + 5]
                    prob += pulp.lpSum(x[(i, d, t, a)] for d in window_days for t in T for a in A) <= 4
        except Exception:
            pass

        for i in I:
            for d in D:
                prob += pulp.lpSum(x[(i, d, t, a)] for t in T for a in A) <= 1

        for d in D:
            for a in A:
                prob += pulp.lpSum(x[(i, d, t, a)] * s.get((i, a), 0) for i in I for t in T) + shortfall_attr[(d, a)] >= n.get((d, a), 0)

        for d in D:
            for t in T:
                prob += pulp.lpSum(x[(i, d, t, a)] for i in I for a in A) + shortfall_pat[(d, t)] >= r.get((d, t), 0)
                prob += pulp.lpSum(x[(i, d, t, a)] for i in I for a in A) <= r.get((d, t), len(I))

        dev_plus = {}
        dev_minus = {}
        for d in D:
            for t in T:
                required = r.get((d, t), 0)
                avg_val = required / max(1, len(A))
                for a in A:
                    dev_plus[(d, t, a)] = pulp.LpVariable(f"dev_plus_{d}_{t}_{a}", lowBound=0, cat="Continuous")
                    dev_minus[(d, t, a)] = pulp.LpVariable(f"dev_minus_{d}_{t}_{a}", lowBound=0, cat="Continuous")
                    attr_count = pulp.lpSum(x[(i, d, t, a)] for i in I)
                    prob += attr_count - avg_val == dev_plus[(d, t, a)] - dev_minus[(d, t, a)]

        P_a, P_t, P_dev = 1000, 500, 50
        prob += (
            pulp.lpSum(P_a * shortfall_attr[(d, a)] for d in D for a in A) +
            pulp.lpSum(P_t * shortfall_pat[(d, t)] for d in D for t in T) +
            pulp.lpSum(P_dev * (dev_plus[(d, t, a)] + dev_minus[(d, t, a)]) for d in D for t in T for a in A)
        )

        solver = pulp.PULP_CBC_CMD(msg=False, timeLimit=300)
        result = prob.solve(solver)
        status = pulp.LpStatus[prob.status]

        if status in ("Optimal", "Feasible"):
            assignment = {}
            for i in I:
                for d in D:
                    for t in T:
                        for a in A:
                            val = pulp.value(x[(i, d, t, a)])
                            if val is not None and val > 0.5:
                                assignment[(i, d)] = (t, a)

            # --- DataFrame生成 ---
            data = []
            for i in I:
                row = []
                for d in D:
                    ta = assignment.get((i, d), ("", ""))
                    row.append(f"{ta[0]}-{ta[1]}" if ta != ("", "") else "")
                data.append(row)
            df_shift = pd.DataFrame(data, index=I, columns=D)

            attribute_rows = []
            for d in D:
                for a in A:
                    required = n.get((d, a), 0)
                    assigned = sum(s.get((i, a), 0) for i in I for t in T if pulp.value(x[(i, d, t, a)]) is not None and pulp.value(x[(i, d, t, a)]) > 0.5)
                    penalty = pulp.value(shortfall_attr[(d, a)]) or 0
                    attribute_rows.append([d, a, required, assigned, penalty])
            df_attribute = pd.DataFrame(attribute_rows, columns=['日付','属性','必要点数','割当点数','不足ペナルティ'])

            pattern_rows = []
            for d in D:
                for t in T:
                    required = r.get((d, t), 0)
                    assigned = sum(1 for i in I for a in A if pulp.value(x[(i, d, t, a)]) is not None and pulp.value(x[(i, d, t, a)]) > 0.5)
                    penalty = pulp.value(shortfall_pat[(d, t)]) or 0
                    pattern_rows.append([d, t, required, assigned, penalty])
            df_pattern = pd.DataFrame(pattern_rows, columns=['日付','勤務パターン','必要人数','割当人数','不足ペナルティ'])

            total_workdays_rows = []
            for i in I:
                total_days = sum(1 for d in D for t in T for a in A if pulp.value(x[(i, d, t, a)]) is not None and pulp.value(x[(i, d, t, a)]) > 0.5)
                total_workdays_rows.append([i, total_days])
            df_total_workdays = pd.DataFrame(total_workdays_rows, columns=['従業員','総勤務日数'])

            dev_rows = []
            for d in D:
                for t in T:
                    required = r.get((d, t), 0)
                    avg = required / max(1, len(A))
                    for a in A:
                        assigned_attr = sum(1 for i in I if pulp.value(x[(i, d, t, a)]) is not None and pulp.value(x[(i, d, t, a)]) > 0.5)
                        dp = pulp.value(dev_plus[(d, t, a)]) or 0
                        dm = pulp.value(dev_minus[(d, t, a)]) or 0
                        dev_rows.append([d, t, a, required, assigned_attr, avg, dp, dm])
            df_dev = pd.DataFrame(dev_rows, columns=['日付','勤務パターン','属性','必要人数','割当人数','平均(必要/属性)','偏り+','偏り-'])

            # --- Excel出力 ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_shift.to_excel(writer, sheet_name='割り当て結果')
                df_attribute.to_excel(writer, sheet_name='属性点数確認', index=False)
                df_pattern.to_excel(writer, sheet_name='パターン人数確認', index=False)
                df_total_workdays.to_excel(writer, sheet_name='勤務日数集計', index=False)
                df_dev.to_excel(writer, sheet_name='属性偏り確認', index=False)
            output.seek(0)

            # 追加: アプリ内表示用にまとめる
            dfs_for_display = {
                "割り当て結果": df_shift,
                "属性点数確認": df_attribute,
                "パターン人数確認": df_pattern,
                "勤務日数集計": df_total_workdays,
                "属性偏り確認": df_dev
            }

            return output, dfs_for_display

        else:
            return None, None

    # ======= Streamlit UI =======
    if uploaded_file:
        st.success("✅ ファイルを読み込みました！")
        if st.button("最適化を実行"):
            with st.spinner("最適化中...（数分かかる場合があります）"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(uploaded_file.read())
                    tmp_path = tmp.name

                output, dfs = run_shift_optimization(tmp_path)

                if output:
                    st.success("✅ 最適化が完了しました！")
                    st.download_button(
                        "📥 結果Excelをダウンロード",
                        data=output.getvalue(),
                        file_name="シフト出力結果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    # 追加: Streamlit上でシート内容確認
                    st.markdown("### 📊 最適化結果確認")
                    for sheet_name, df in dfs.items():
                        st.subheader(sheet_name)
                        st.dataframe(df)

                else:
                    st.error("❌ 解が見つかりませんでした。制約条件を確認してください。")
                os.remove(tmp_path)
    else:
        st.info("⬆️ 入力表をアップロードしてください。")
