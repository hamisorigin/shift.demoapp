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
    st.title("⚙️ シフト自動作成（MIP版）")

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


        model = Model("ShiftScheduling_4D_Penalty_Dev_NoOverstaff")
        x = {(i, d, t, a): model.add_var(var_type=BINARY) for i in I for d in D for t in T for a in A}
        shortfall_attr = {(d, a): model.add_var(lb=0) for d in D for a in A}
        shortfall_pat = {(d, t): model.add_var(lb=0) for d in D for t in T}

        for i in I:
            for d in D:
                for t in T:
                    for a in A:
                        model += x[i, d, t, a] <= k[i, d]
                        model += x[i, d, t, a] <= g[i, t]
            model += xsum(x[i, d, t, a] for d in D for t in T for a in A) >= l_min[i]
            model += xsum(x[i, d, t, a] for d in D for t in T for a in A) <= l_max[i]

        try:
            D_numeric = sorted([int(d) for d in D])
            for i in I:
                for idx in range(len(D_numeric) - 4):
                    window_days = D_numeric[idx:idx + 5]
                    model += xsum(x[i, d, t, a] for d in window_days for t in T for a in A) <= 4
        except:
            pass

        for i in I:
            for d in D:
                model += xsum(x[i, d, t, a] for t in T for a in A) <= 1

        for d in D:
            for a in A:
                model += xsum(x[i, d, t, a]*s[i, a] for i in I for t in T) + shortfall_attr[d, a] >= n[d, a]
            for t in T:
                model += xsum(x[i, d, t, a] for i in I for a in A) + shortfall_pat[d, t] >= r[d, t]
                model += xsum(x[i, d, t, a] for i in I for a in A) <= r[d, t]

        dev_plus, dev_minus = {}, {}
        for d in D:
            for t in T:
                required = r[d, t]
                avg_val = required / max(1, len(A))
                for a in A:
                    attr_count = xsum(x[i, d, t, a] for i in I)
                    dev_plus[d, t, a] = model.add_var(lb=0)
                    dev_minus[d, t, a] = model.add_var(lb=0)
                    model += attr_count - avg_val == dev_plus[d, t, a] - dev_minus[d, t, a]

        P_a, P_t, P_dev = 1000, 500, 50
        model.objective = minimize(
            xsum(P_a * shortfall_attr[d, a] for d in D for a in A) +
            xsum(P_t * shortfall_pat[d, t] for d in D for t in T) +
            xsum(P_dev * (dev_plus[d, t, a] + dev_minus[d, t, a]) for d in D for t in T for a in A)
        )

        status = model.optimize()

        # === 出力 ===
        if status in [OptimizationStatus.OPTIMAL, OptimizationStatus.FEASIBLE]:
            assignment = {}
            for i in I:
                for d in D:
                    for t in T:
                        for a in A:
                            if x[i, d, t, a].x > 0.5:
                                assignment[(i, d)] = (t, a)

            data = []
            for i in I:
                row = []
                for d in D:
                    ta = assignment.get((i, d), ("", ""))
                    row.append(f"{ta[0]}-{ta[1]}" if ta != ("", "") else "")
                data.append(row)
            df_shift = pd.DataFrame(data, index=I, columns=D)

            attribute_rows, pattern_rows, total_workdays_rows, dev_rows = [], [], [], []

            for d in D:
                for a in A:
                    required = n.get((d, a), 0)
                    assigned = sum(s[i, a] for i in I for t in T if x[i, d, t, a].x > 0.5)
                    penalty = shortfall_attr[d, a].x or 0
                    attribute_rows.append([d, a, required, assigned, penalty])
            df_attribute = pd.DataFrame(attribute_rows, columns=['日付','属性','必要点数','割当点数','不足ペナルティ'])

            for d in D:
                for t in T:
                    required = r[d, t]
                    assigned = sum(1 for i in I for a in A if x[i, d, t, a].x > 0.5)
                    penalty = shortfall_pat[d, t].x or 0
                    pattern_rows.append([d, t, required, assigned, penalty])
            df_pattern = pd.DataFrame(pattern_rows, columns=['日付','勤務パターン','必要人数','割当人数','不足ペナルティ'])

            for i in I:
                total_days = sum(1 for d in D for t in T for a in A if x[i, d, t, a].x > 0.5)
                total_workdays_rows.append([i, total_days])
            df_total_workdays = pd.DataFrame(total_workdays_rows, columns=['従業員','総勤務日数'])

            for d in D:
                for t in T:
                    required = r[d, t]
                    avg = required / max(1, len(A))
                    for a in A:
                        assigned_attr = sum(1 for i in I if x[i, d, t, a].x > 0.5)
                        dp = dev_plus[d, t, a].x or 0
                        dm = dev_minus[d, t, a].x or 0
                        dev_rows.append([d, t, a, required, assigned_attr, avg, dp, dm])
            df_dev = pd.DataFrame(dev_rows, columns=['日付','勤務パターン','属性','必要人数','割当人数','平均(必要/属性)','偏り+','偏り-'])

            # === Excel出力 ===
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_shift.to_excel(writer, sheet_name='割り当て結果')
                df_attribute.to_excel(writer, sheet_name='属性点数確認', index=False)
                df_pattern.to_excel(writer, sheet_name='パターン人数確認', index=False)
                df_total_workdays.to_excel(writer, sheet_name='勤務日数集計', index=False)
                df_dev.to_excel(writer, sheet_name='属性偏り確認', index=False)
            output.seek(0)
            return output
        else:
            return None

    # ======= Streamlit UI =======
    if uploaded_file:
        st.success("✅ ファイルを読み込みました！")
        if st.button("最適化を実行"):
            with st.spinner("最適化中...（数分かかる場合があります）"):
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(uploaded_file.read())
                    tmp_path = tmp.name

                output = run_shift_optimization(tmp_path)

                if output:
                    st.success("✅ 最適化が完了しました！")
                    st.download_button(
                        "📥 結果Excelをダウンロード",
                        data=output.getvalue(),
                        file_name="シフト出力結果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("❌ 解が見つかりませんでした。制約条件を確認してください。")
                os.remove(tmp_path)
    else:
        st.info("⬆️ 入力表をアップロードしてください。")
