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

page = st.sidebar.radio("ページを選択", ["テンプレート作成", "シフト最適化"])

# ==============================
# ページ① テンプレート作成
# ==============================
if page == "テンプレート作成":
    st.title("📋 シフト入力テンプレート自動生成（下限/上限対応版）")

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
    D = [i + 1 for i in range(num_days)]

    if st.button("📄 テンプレートExcelを生成"):
        df_availability = pd.DataFrame("", index=I, columns=D)
        df_availability.index.name = "従業員"

        df_pattern = pd.DataFrame("", index=I, columns=T)
        df_pattern.index.name = "従業員"

        df_limits = pd.DataFrame({"従業員": I, "下限": [0]*len(I), "上限": [num_days]*len(I)})

        df_ability = pd.DataFrame("", index=I, columns=A)
        df_ability.index.name = "従業員"

        df_need_attr = pd.DataFrame("", index=D, columns=A)
        df_need_attr.index.name = "日付"

        # ✅ 縦形式の必要勤務人数
        df_need_pattern_bounds = pd.DataFrame(
            [[d, t, 0, 0] for d in D for t in T],
            columns=["日付", "出勤パターン", "下限", "上限"]
        )

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_availability.to_excel(writer, sheet_name='出勤可能日')
            df_pattern.to_excel(writer, sheet_name='勤務可能パターン')
            df_limits.to_excel(writer, sheet_name='勤務日数上下限', index=False)
            df_ability.to_excel(writer, sheet_name='従業員能力表')
            df_need_attr.to_excel(writer, sheet_name='属性ごとの必要点数')
            df_need_pattern_bounds.to_excel(writer, sheet_name='必要勤務人数', index=False)

        st.download_button(
            label="📥 テンプレートExcelをダウンロード",
            data=output.getvalue(),
            file_name="シフト自動作成汎用アプリ-入力表-下限上限対応.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==============================
# ページ② シフト最適化
# ==============================
elif page == "シフト最適化":
    st.title("⚙️ シフト自動作成（PuLP版） — 下限ハード / 上限ソフト / 属性偏り最小化")

    uploaded_file = st.file_uploader("📤 Excelファイルをアップロード", type=["xlsx"])

    # ✅ ベテラン制約を適用する勤務パターンを入力（空白なら無効）
    st.markdown("※ 従業員能力値は **1〜10 の範囲で入力してください（10 がベテラン）**")
    veteran_pattern = st.text_input(
        "ベテラン（能力10の従業員）を最低1人配置したい勤務パターン（例：遅番）",
        value="",
        help="特定の勤務パターンに能力10の人を必ず1人入れたい場合に入力（空白なら無効）"
    )

    def run_shift_optimization(file_path):
        filename = file_path

        # 共通読み込み関数
        def extract_sheet_data_generic(file, sheet_name):
            try:
                df = pd.read_excel(file, sheet_name=sheet_name, header=None)
            except Exception:
                return []
            row_labels = df.iloc[1:, 0].dropna().tolist()
            col_labels = df.iloc[0, 1:].dropna().tolist()
            result = []
            for i, row in enumerate(row_labels, start=1):
                for j, col in enumerate(col_labels, start=1):
                    val = df.iat[i, j]
                    if pd.notna(val):
                        result.append((row, col, float(val)))
            return result

        # --- 各シート読み込み ---
        availability_list = extract_sheet_data_generic(filename, '出勤可能日')
        pattern_list = extract_sheet_data_generic(filename, '勤務可能パターン')
        employeeability_list = extract_sheet_data_generic(filename, '従業員能力表')
        needwork_list = extract_sheet_data_generic(filename, '属性ごとの必要点数')
        df_limits = pd.read_excel(filename, sheet_name='勤務日数上下限')

        l_min = dict(zip(df_limits['従業員'], df_limits['下限']))
        l_max = dict(zip(df_limits['従業員'], df_limits['上限']))

        I = sorted(set([r[0] for r in availability_list]))
        D = sorted(set([r[1] for r in availability_list]))
        T = sorted(set([r[1] for r in pattern_list]))
        A = sorted(set([r[1] for r in employeeability_list]))

        # ✅ 必要勤務人数（縦形式）
        r_min, r_max = {}, {}
        df_req = pd.read_excel(filename, sheet_name='必要勤務人数')
        for _, row in df_req.iterrows():
            d = int(row['日付'])
            t = str(row['出勤パターン']).strip()
            r_min[(d, t)] = int(row['下限']) if pd.notna(row['下限']) else 0
            r_max[(d, t)] = int(row['上限']) if pd.notna(row['上限']) else len(I)

        # --- 辞書整形 ---
        k = {(i, d): 0 for i in I for d in D}
        for i, d, val in availability_list:
            k[i, d] = int(val)

        g = {(i, t): 0 for i in I for t in T}
        for i, t, val in pattern_list:
            g[i, t] = int(val)

        s = {(i, a): 0 for i in I for a in A}
        for i, a, val in employeeability_list:
            s[i, a] = float(val)

        n = {(d, a): 0 for d in D for a in A}
        for d, a, val in needwork_list:
            n[d, a] = float(val)

        # --- モデル ---
        prob = pulp.LpProblem("ShiftScheduling", pulp.LpMinimize)
        x = pulp.LpVariable.dicts("x", (I, D, T, A), 0, 1, cat="Binary")
        short_a = pulp.LpVariable.dicts("short_attr", (D, A), 0)
        over_t = pulp.LpVariable.dicts("over_pat", (D, T), 0)

        # --- 制約 ---
        # 出勤・パターン制約
        for i in I:
            for d in D:
                for t in T:
                    for a in A:
                        prob += x[i][d][t][a] <= k[i, d]
                        prob += x[i][d][t][a] <= g[i, t]
                        if s[i, a] == 0:
                            prob += x[i][d][t][a] == 0

        # 勤務日数制約
        for i in I:
            prob += pulp.lpSum(x[i][d][t][a] for d in D for t in T for a in A) >= l_min[i]
            prob += pulp.lpSum(x[i][d][t][a] for d in D for t in T for a in A) <= l_max[i]

        # --- 5連勤防止制約 ---
        D_numeric = sorted([int(d) for d in D if str(d).isdigit()])
        for i in I:
            for idx in range(len(D_numeric) - 4):
                window_days = D_numeric[idx:idx + 5]
                prob += pulp.lpSum(x[i][d][t][a] for d in window_days for t in T for a in A) <= 4

        # 1日1勤務
        for i in I:
            for d in D:
                prob += pulp.lpSum(x[i][d][t][a] for t in T for a in A) <= 1

        # 属性点数制約
        for d in D:
            for a in A:
                prob += pulp.lpSum(x[i][d][t][a] * s[i, a] for i in I for t in T) + short_a[d][a] >= n[d, a]

        # パターン人数制約
        for d in D:
            for t in T:
                prob += pulp.lpSum(x[i][d][t][a] for i in I for a in A) >= r_min[(d, t)]
                prob += pulp.lpSum(x[i][d][t][a] for i in I for a in A) - over_t[d][t] <= r_max[(d, t)]


        # ✅ ベテラン制約（ユーザー指定）
        if veteran_pattern.strip() != "":
            if veteran_pattern not in T:
                st.warning(f"⚠️ 勤務パターン「{veteran_pattern}」は存在しません。ベテラン制約は無効です。")
            else:
                st.info(f"🧩 ベテラン制約を適用中：『{veteran_pattern}』に能力10の人を最低1人配置")
                for d in D:
                    for t in T:
                        if t == veteran_pattern:
                            for a in A:
                                capable_workers = [
                                    i for i in I if s[i, a] == 10 and k[i, d] == 1 and g[i, t] == 1
                                ]
                                if capable_workers:
                                    prob += pulp.lpSum(x[i][d][t][a] for i in capable_workers) >= 1
        else:
            st.info("🧩 ベテラン制約は適用されません（入力なし）")

        # ✅ 属性偏り制約（復活）
        dev_plus, dev_minus = {}, {}
        for d in D:
            for t in T:
                required = r_min.get((d, t), 0)
                avg_val = required / max(1, len(A))
                for a in A:
                    dev_plus[d, t, a] = pulp.LpVariable(f"dev_plus_{d}_{t}_{a}", lowBound=0)
                    dev_minus[d, t, a] = pulp.LpVariable(f"dev_minus_{d}_{t}_{a}", lowBound=0)
                    attr_count = pulp.lpSum(x[i][d][t][a] for i in I)
                    prob += attr_count - avg_val == dev_plus[d, t, a] - dev_minus[d, t, a]

        # ✅ 目的関数
        P_short_a, P_over_t, P_dev = 1000, 200, 50
        prob += (
            pulp.lpSum(P_short_a * short_a[d][a] for d in D for a in A)
            + pulp.lpSum(P_over_t * over_t[d][t] for d in D for t in T)
            + pulp.lpSum(P_dev * (dev_plus[d, t, a] + dev_minus[d, t, a]) for d in D for t in T for a in A)
        )

        # --- ソルバー実行 ---
        solver = pulp.PULP_CBC_CMD(msg=False)
        prob.solve(solver)


        # --- ペナルティ集計 ---
        penalty_short = sum(pulp.value(short_a[d][a]) for d in D for a in A)
        penalty_over = sum(pulp.value(over_t[d][t]) for d in D for t in T)
        penalty_dev = sum(pulp.value(dev_plus[d, t, a]) + pulp.value(dev_minus[d, t, a]) for d in D for t in T for a in A)

        total_penalty = (
            200 * penalty_short +
            100 * penalty_over +
            50 * penalty_dev
        )

        # --- Streamlit表示部分 ---
        st.subheader("📊 ペナルティ集計結果")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("属性不足ペナルティ", f"{penalty_short:.1f}")
        with col2:
            st.metric("人数超過ペナルティ", f"{penalty_over:.1f}")
        with col3:
            st.metric("偏りペナルティ", f"{penalty_dev:.1f}")
        with col4:
            st.metric("総合ペナルティスコア", f"{total_penalty:.1f}")

        # --- 出力整形 ---
        assignment = {(i, d): "" for i in I for d in D}
        for i in I:
            for d in D:
                for t in T:
                    for a in A:
                        if pulp.value(x[i][d][t][a]) > 0.5:
                            assignment[(i, d)] = f"{t}-{a}"

        df_shift = pd.DataFrame([[assignment[(i, d)] for d in D] for i in I], index=I, columns=D)

        # 勤務日数集計
        df_days = pd.DataFrame([
            [i, sum(1 for d in D for t in T for a in A if pulp.value(x[i][d][t][a]) > 0.5), l_min[i], l_max[i]]
            for i in I
        ], columns=["従業員", "総勤務日数", "下限", "上限"])
        df_days["判定"] = df_days.apply(
            lambda r: "不足" if r["総勤務日数"] < r["下限"] else ("超過" if r["総勤務日数"] > r["上限"] else "OK"), axis=1
        )

        # 属性点数確認
        df_attr = pd.DataFrame([
            [d, a, n[d, a],
             sum(s[i, a] for i in I for t in T if pulp.value(x[i][d][t][a]) > 0.5),
             pulp.value(short_a[d][a])]
            for d in D for a in A
        ], columns=["日付", "属性", "必要点数", "割当点数", "不足ペナルティ"])

        # パターン人数確認
        df_pattern = pd.DataFrame([
            [d, t, r_min[(d, t)], r_max[(d, t)],
             sum(1 for i in I for a in A if pulp.value(x[i][d][t][a]) > 0.5),
             pulp.value(over_t[d][t])]
            for d in D for t in T
        ], columns=["日付", "勤務パターン", "下限", "上限", "割当人数", "上限超過ペナルティ"])

        # 属性偏り確認
        df_dev = pd.DataFrame([
            [d, t, a,
             r_min.get((d, t), 0),
             sum(1 for i in I if pulp.value(x[i][d][t][a]) > 0.5),
             r_min.get((d, t), 0)/max(1, len(A)),
             pulp.value(dev_plus[d, t, a]),
             pulp.value(dev_minus[d, t, a])]
            for d in D for t in T for a in A
        ], columns=["日付", "勤務パターン", "属性", "必要人数", "割当人数", "平均(必要/属性)", "偏り+", "偏り-"])

        # Excel出力
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_shift.to_excel(writer, sheet_name="割り当て結果")
            df_days.to_excel(writer, sheet_name="勤務日数集計", index=False)
            df_attr.to_excel(writer, sheet_name="属性点数確認", index=False)
            df_pattern.to_excel(writer, sheet_name="パターン人数確認", index=False)
            df_dev.to_excel(writer, sheet_name="属性偏り確認", index=False)

        output.seek(0)
        dfs = {
            "割り当て結果": df_shift,
            "勤務日数集計": df_days,
            "属性点数確認": df_attr,
            "パターン人数確認": df_pattern,
            "属性偏り確認": df_dev
        }
        return output, dfs

    # --- UI ---
    if uploaded_file:
        if st.button("最適化を実行"):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(uploaded_file.read())
                tmp_path = tmp.name

            output, dfs = run_shift_optimization(tmp_path)
            if output:
                st.success("✅ 最適化完了！")
                st.download_button("📥 結果Excelをダウンロード",
                                   data=output.getvalue(),
                                   file_name="シフト出力結果.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                for k, v in dfs.items():
                    st.subheader(k)
                    st.dataframe(v)
            os.remove(tmp_path)
