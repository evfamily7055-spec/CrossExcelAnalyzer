import streamlit as st
import pandas as pd
import io
import plotly.express as px
import statsmodels.api as sm
from scipy import stats
import plotly.graph_objects as go
import plotly.io as pio # 画像エクスポートのために追加

# --- ページ設定 (Page Config) ---
st.set_page_config(layout="wide")
st.title("Excelデータアナライザー 📈")
st.info("Excelファイルをアップロードすると、①データの自動概要、②ピボット分析、③統計解析 を実行します。")

# --- セッションステートの初期化 (Initialize Session State) ---
if 'pivot_df' not in st.session_state:
    st.session_state.pivot_df = None
if 'pivot_config' not in st.session_state:
    st.session_state.pivot_config = {}
if 'df' not in st.session_state:
    st.session_state.df = None

# --- 1. ファイルアップローダー (File Uploader) ---
uploaded_file = st.file_uploader("Excelファイル (.xlsx) をアップロードしてください", type=["xlsx"])

if uploaded_file:
    try:
        bytes_data = uploaded_file.getvalue()
        df = pd.read_excel(io.BytesIO(bytes_data))
        st.session_state.df = df # 完全なデータフレームをセッションステートに保存
        
        st.success("ファイルの読み込みが完了しました。")

    except Exception as e:
        st.error(f"Excelファイルの読み込みに失敗しました: {e}")
        st.session_state.df = None
else:
    st.info("👆 上のボタンからExcelファイルをアップロードして開始します。")


# --- メインアプリロジック (データフレーム読み込み後) ---
if st.session_state.df is not None:
    df = st.session_state.df

    # --- (UX-IMPROVED v4) 1. データの自動概要（速報値） ---
    st.markdown("---")
    st.header("1. データの自動概要（速報値）")
    st.write("各項目（列）のデータ分布を自動で可視化します。")
    st.info("グラフにマウスを合わせると右上に表示される **カメラアイコン** からも画像を保存できます。")
    
    # 全体の基本情報
    col1, col2, col3 = st.columns(3)
    col1.metric("総回答数（行）", f"{len(df):,}")
    col2.metric("総項目数（列）", f"{len(df.columns)}")
    col3.metric("欠損値の合計", f"{df.isnull().sum().sum():,}")
    
    # (NEW) 要約情報をためておくリストを初期化
    excluded_cols_summary = []
    numeric_summary_list = [] # 数値データの統計量をここに蓄積
    
    # カラーパレットの定義
    color_palette_pie = px.colors.qualitative.Pastel
    color_palette_bar = px.colors.qualitative.Safe
    
    # 「その他」に丸める際の閾値
    TOP_N_FOR_OTHERS = 9 # 上位9位まで表示
    OTHERS_THRESHOLD = 0.2 # 「その他」が20%未満の場合に円グラフ採用

    # --- グラフ描画ループ ---
    st.subheader("📊 各項目のグラフ一覧")
    for col in df.columns:
        st.markdown(f"--- \n ### 項目: {col}") # H3レベルに変更
        
        # 1. 数値データ (Numeric Data) の場合 -> ヒストグラム
        if pd.api.types.is_numeric_dtype(df[col]):
            st.write(f"（数値データとして認識）")
            
            # (NEW) 統計量を計算し、リストに追加（表示はしない）
            try:
                desc_stats = df[col].describe()
                numeric_summary_list.append({
                    "項目名": col,
                    "平均値": desc_stats['mean'],
                    "中央値": desc_stats['50%'],
                    "最小値": desc_stats['min'],
                    "最大値": desc_stats['max'],
                    "標準偏差": desc_stats['std'],
                    "件数": desc_stats['count']
                })
            except Exception:
                pass # 統計計算失敗時はスキップ

            # ヒストグラムを描画
            try:
                fig = px.histogram(
                    df, 
                    x=col, 
                    title=f"「{col}」の分布（ヒストグラム）",
                    marginal="box"
                )
                st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.warning(f"グラフ描画失敗 (数値): {e}")

        # 2. カテゴリデータ (Categorical Data) の場合
        elif pd.api.types.is_string_dtype(df[col]) or pd.api.types.is_object_dtype(df[col]):
            st.write(f"（カテゴリ/テキストデータとして認識）")
            
            n_unique = df[col].nunique()
            counts = df[col].value_counts()
            
            # (A) ユニーク数が 2〜10 の場合 (選択式設問と推定) -> 円グラフ
            if 1 < n_unique <= 10:
                try:
                    df_pie = counts.reset_index()
                    df_pie.columns = ['value', 'count']
                    
                    fig = px.pie(
                        df_pie, 
                        names='value', 
                        values='count', 
                        title=f"「{col}」の構成比（円グラフ）",
                        color_discrete_sequence=color_palette_pie
                    )
                    fig.update_traces(textposition='inside', textinfo='percent+label')
                    st.plotly_chart(fig, use_container_width=True)
                except Exception as e:
                    st.warning(f"グラフ描画失敗 (円グラフ): {e}")

            # (B) ユニーク数が 11〜20 の場合 -> 「その他」ロジックを試行
            elif 10 < n_unique <= 20:
                try:
                    total_count = counts.sum()
                    top_n_counts = counts.iloc[:TOP_N_FOR_OTHERS]
                    other_count = counts.iloc[TOP_N_FOR_OTHERS:].sum()
                    other_percentage = other_count / total_count
                    
                    if other_count > 0 and other_percentage < OTHERS_THRESHOLD:
                        # (B-1) 「その他」が閾値未満なら、丸めて円グラフ
                        st.write(f"（上位{TOP_N_FOR_OTHERS}件と「その他」で表示）")
                        other_row = pd.Series([other_count], index=['その他'])
                        df_pie_data = pd.concat([top_n_counts, other_row])
                        
                        df_pie = df_pie_data.reset_index()
                        df_pie.columns = ['value', 'count']

                        fig = px.pie(
                            df_pie, 
                            names='value', 
                            values='count', 
                            title=f"「{col}」の構成比（円グラフ, その他集約）",
                            color_discrete_sequence=color_palette_pie
                        )
                        fig.update_traces(textposition='inside', textinfo='percent+label')
                        st.plotly_chart(fig, use_container_width=True)
                    
                    else:
                        # (B-2) 「その他」が閾値以上なら、丸めずに棒グラフ
                        st.write("（回答が分散しているため棒グラフで表示）")
                        df_bar = counts.reset_index()
                        df_bar.columns = ['value', 'count']
                        
                        fig = px.bar(
                            df_bar, 
                            x='value', 
                            y='count',
                            title=f"「{col}」の内訳（棒グラフ）",
                            text='count',
                            color='value',
                            color_discrete_sequence=color_palette_bar
                        )
                        fig.update_xaxes(title_text=col)
                        st.plotly_chart(fig, use_container_width=True)
                        
                except Exception as e:
                    st.warning(f"グラフ描画失敗 (カテゴリ/その他): {e}")
            
            # (C) ユニーク数が多すぎる(>20) or 少なすぎる(<=1)場合
            else:
                if n_unique > 20:
                    reason = f"回答の種類が {n_unique} と多すぎるため（フリーテキストの可能性）"
                elif n_unique <= 1:
                    reason = f"回答の種類が1種類以下のため"
                
                # (NEW) 警告を表示せず、リストに追加のみ
                excluded_cols_summary.append((col, reason))
        
        # 3. 日付データ (Datetime Data) の場合 -> 折れ線グラフ
        elif pd.api.types.is_datetime64_any_dtype(df[col]):
            st.write(f"（日付データとして認識）")
            try:
                counts_over_time = df[col].dt.date.value_counts().sort_index().reset_index()
                counts_over_time.columns = ['date', 'count']
                
                fig = px.line(
                    counts_over_time,
                    x='date',
                    y='count',
                    title=f"「{col}」の時系列（回答件数）",
                    markers=True
                )
                st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.warning(f"グラフ描画失敗 (日付): {e}")
        
        # 4. その他のデータ型
        else:
            reason = f"認識できないデータ型 ({df[col].dtype}) のため"
            excluded_cols_summary.append((col, reason))

    # --- (NEW) ループ終了後、要約情報をまとめて表示 ---
    st.markdown("---")
    st.subheader("📝 数値データの基本統計量まとめ")
    
    if numeric_summary_list:
        summary_df = pd.DataFrame(numeric_summary_list).set_index("項目名")
        # 数値を見やすくフォーマット
        st.dataframe(summary_df.style.format("{:,.2f}"), use_container_width=True)
    else:
        st.info("集計対象となる数値データがありませんでした。")

    # (NEW) スキップした項目も最後にまとめて表示
    if excluded_cols_summary:
        st.markdown("---")
        st.subheader("🚫 集計から外した項目")
        st.write("以下の項目は、回答の種類が多すぎる（フリーテキスト等）か、データ型が特殊なため自動概要のグラフ化から除外されました。")
        
        excluded_df = pd.DataFrame(excluded_cols_summary, columns=["項目名 (ラベル名)", "除外理由"])
        st.dataframe(excluded_df, use_container_width=True)


    # --- 2. データプレビュー (Data Preview) ---
    st.markdown("---")
    st.header("2. データプレビュー (先頭10行)")
    with st.expander("データ全体を表示する"):
        st.dataframe(df.head(10), use_container_width=True)

    # --- 3. ピボットテーブル & グラフセクション (Pivot Table & Graph Section) ---
    st.markdown("---")
    st.header("3. ピボットテーブル & グラフ")
    
    with st.expander("ピボットテーブルの分析はこちらをクリック"):
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("ピボット設定")
            
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            
            index_cols = st.multiselect(
                "行 (Rows)", options=df.columns.tolist(), key="pivot_index"
            )
            column_cols = st.multiselect(
                "列 (Columns)", options=df.columns.tolist(), key="pivot_cols"
            )
            value_col = st.selectbox(
                "値 (Values)", options=numeric_cols, index=None, placeholder="集計する数値列...", key="pivot_val"
            )
            agg_func_options = {"合計": "sum", "平均": "mean", "カウント": "count", "中央値": "median"}
            agg_func_label = st.selectbox(
                "集計方法", options=agg_func_options.keys(), key="pivot_agg"
            )
            agg_func = agg_func_options.get(agg_func_label, "sum")

            if st.button("ピボットテーブル作成", type="primary"):
                if index_cols and value_col:
                    try:
                        pivot_df = pd.pivot_table(
                            df,
                            index=index_cols,
                            columns=column_cols if column_cols else None,
                            values=value_col,
                            aggfunc=agg_func,
                            fill_value=0 
                        )
                        st.session_state.pivot_df = pivot_df
                        st.session_state.pivot_config = {
                            "index": index_cols, "columns": column_cols, "values": value_col, "agg_label": agg_func_label
                        }
                    except Exception as e:
                        st.error(f"ピボット作成失敗: {e}")
                        st.session_state.pivot_df = None
                else:
                    st.warning("「行」と「値」は必須です。")

        with col2:
            if st.session_state.pivot_df is not None:
                st.subheader("ピボット実行結果")
                st.dataframe(st.session_state.pivot_df, use_container_width=True)
                
                @st.cache_data
                def convert_df_to_csv(df_to_convert):
                    return df_to_convert.to_csv(index=True).encode('utf-8-sig')
                
                csv_data = convert_df_to_csv(st.session_state.pivot_df)
                st.download_button(label="結果をCSVダウンロード", data=csv_data, file_name="pivot.csv", mime="text/csv")

        # --- ピボットグラフの可視化 (Pivot Graph Visualization) ---
        if st.session_state.pivot_df is not None:
            st.subheader("ピボットグラフ可視化")
            
            pivot_df = st.session_state.pivot_df
            config = st.session_state.pivot_config

            chart_type = st.selectbox(
                "グラフの種類を選択",
                options=[
                    "ヒートマップ", "グループ棒グラフ", "積み上げ棒グラフ", "折れ線グラフ", "円グラフ"
                ],
                key="pivot_chart_type"
            )

            try:
                fig_pivot = None # グラフオブジェクトを格納する変数を初期化
                if chart_type == "円グラフ":
                    if len(config["index"]) == 1 and not config["columns"]:
                        df_for_pie = pivot_df.reset_index()
                        names_col = config["index"][0]
                        values_col = config["values"]
                        fig_pivot = px.pie(df_for_pie, names=names_col, values=values_col, 
                                     title=f"{values_col} ({config['agg_label']}) の構成比",
                                     color_discrete_sequence=color_palette_pie)
                        fig_pivot.update_traces(textposition='inside', textinfo='percent+label')
                    else:
                        st.warning("円グラフは「行」が1項目かつ「列」が空の場合のみ描画されます。")
                
                else:
                    if chart_type == "ヒートマップ":
                        fig_pivot = px.imshow(pivot_df, text_auto=True, aspect="auto", title=f"{config['values']} ({config['agg_label']}) ヒートマップ")
                    elif chart_type == "グループ棒グラフ":
                        fig_pivot = px.bar(pivot_df, barmode='group', title=f"{config['values']} ({config['agg_label']}) グループ棒グラフ",
                                     color_discrete_sequence=color_palette_bar)
                    elif chart_type == "積み上げ棒グラフ":
                        fig_pivot = px.bar(pivot_df, barmode='stack', title=f"{config['values']} ({config['agg_label']}) 積み上げ棒グラフ",
                                     color_discrete_sequence=color_palette_bar)
                    elif chart_type == "折れ線グラフ":
                        fig_pivot = px.line(pivot_df, title=f"{config['values']} ({config['agg_label']}) 折れ線グラフ",
                                      markers=True, color_discrete_sequence=color_palette_bar)
                    
                if fig_pivot:
                    st.plotly_chart(fig_pivot, use_container_width=True)
                    
                    img_bytes = pio.to_image(fig_pivot, format="png", scale=2)
                    st.download_button(
                        label="📈 グラフを画像(PNG)で保存",
                        data=img_bytes,
                        file_name=f"pivot_chart_{chart_type}.png",
                        mime="image/png"
                    )
                    
            except Exception as e:
                st.error(f"グラフ描画失敗: {e}")


    # --- 4. 統計解析セクション (Statistical Analysis Section) ---
    st.markdown("---")
    st.header("4. 統計解析")
    
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    categorical_cols = df.select_dtypes(exclude=['number']).columns.tolist()
    
    analysis_type = st.selectbox(
        "実行する分析を選択",
        options=["---", "単回帰分析 (Linear Regression)", "t検定 (Independent t-test)"],
        key="analysis_select"
    )

    # --- B-1: 単回帰分析 (Linear Regression) ---
    if analysis_type == "単回帰分析 (Linear Regression)":
        st.subheader("単回帰分析")
        st.write("2つの数値変数の関係性を分析します（Y = aX + b）。")
        
        col1, col2 = st.columns([1, 2])
        
        with col1:
            y_var = st.selectbox("目的変数 (Y)", options=numeric_cols, index=None, help="予測したい変数（結果）", key="reg_y")
            x_var = st.selectbox("説明変数 (X)", options=numeric_cols, index=None, help="予測に使う変数（原因）", key="reg_x")
            
            fig_reg = None

        if y_var and x_var:
            try:
                fig_reg = px.scatter(
                    df, x=x_var, y=y_var, 
                    trendline="ols", 
                    title=f"{y_var} vs {x_var} の散布図と回帰直線"
                )
                with col2:
                    st.plotly_chart(fig_reg, use_container_width=True)
                    
                    img_bytes_reg = pio.to_image(fig_reg, format="png", scale=2)
                    st.download_button(
                        label="📈 散布図を画像(PNG)で保存",
                        data=img_bytes_reg,
                        file_name=f"regression_{y_var}_vs_{x_var}.png",
                        mime="image/png"
                    )
                
                X = sm.add_constant(df[x_var].dropna())
                Y = df[y_var]
                model = sm.OLS(Y, X, missing='drop').fit()
                
                st.subheader("回帰分析の結果")
                st.metric("決定係数 (R-squared)", f"{model.rsquared:.4f}")
                st.write(f"（{y_var} の変動の {model.rsquared:.1%} が {x_var} で説明可能です）")
                st.code(f"{y_var} = {model.params[x_var]:.4f} * {x_var} + {model.params['const']:.4f}")
                st.text(model.summary())

            except Exception as e:
                st.error(f"回帰分析 実行エラー: {e}")

    # --- B-2: t検定 (t-test) ---
    elif analysis_type == "t検定 (Independent t-test)":
        st.subheader("t検定（独立2群間の平均値の差）")
        
        col1, col2 = st.columns([1, 2])
        
        with col1:
            numeric_var = st.selectbox("比較したい数値列", options=numeric_cols, index=None, key="ttest_num")
            
            suitable_cat_cols = [col for col in categorical_cols if 2 <= df[col].nunique() < 20]
            other_cat_cols = [col for col in categorical_cols if col not in suitable_cat_cols]
            sorted_cat_cols = suitable_cat_cols + other_cat_cols
            
            group_var = st.selectbox("グループ分け列", options=sorted_cat_cols, index=None, key="ttest_group")
            
            group1_val, group2_val = None, None
            if group_var:
                unique_groups = df[group_var].dropna().unique()
                if len(unique_groups) == 2:
                    st.success(f"'{unique_groups[0]}' と '{unique_groups[1]}' の2群を比較します。")
                    group1_val, group2_val = unique_groups[0], unique_groups[1]
                else:
                    st.warning(f"グループが {len(unique_groups)} 個あります。比較したい2つを選んでください。")
                    selected_groups = st.multiselect("比較する2つのグループを選択", options=unique_groups, key="ttest_groups_select")
                    if len(selected_groups) == 2:
                        group1_val, group2_val = selected_groups[0], selected_groups[1]

        if numeric_var and group_var and group1_val and group2_val:
            try:
                group1_data = df[df[group_var] == group1_val][numeric_var].dropna()
                group2_data = df[df[group_var] == group2_val][numeric_var].dropna()
                
                t_stat, p_value = stats.ttest_ind(group1_data, group2_data, equal_var=False)
                
                with col2:
                    st.subheader(f"'{group1_val}' vs '{group2_val}' の平均値比較")
                    mean1, mean2 = group1_data.mean(), group2_data.mean()
                    st.metric(f"平均値: {group1_val}", f"{mean1:.4f} (n={len(group1_data)})")
                    st.metric(f"平均値: {group2_val}", f"{mean2:.4f} (n={len(group2_data)})")

                    st.subheader("t検定の結果")
                    st.metric("p値 (p-value)", f"{p_value:.4f}")
                    if p_value < 0.05:
                        st.success("p値 < 0.05: 2群の平均値に統計的に有意な差があると言えます。")
                    else:
                        st.warning("p値 >= 0.05: 2群の平均値に統計的な差があるとは言えません。")
                
            except Exception as e:
                st.error(f"t検定 実行エラー: {e}")

# --- フッター (ファイル未読み込み時) (Footer) ---
elif not uploaded_file:
    st.markdown("---")
    st.header("利用可能な分析")
    st.write("1. **データの自動概要**:（Google Forms風）アップロードされたデータの全項目を自動で集計・グラフ化します。")
    st.write("2. **ピボットテーブル & グラフ**: Excelライクなクロス集計と、円グラフや棒グラフによる可視化。")
    st.write("3. **単回帰分析**: 2つの数値データ（例：広告費と売上）の関係性を分析し、傾向線を表示。")
    st.write("4. **t検定**: 2つのグループ（例：AプランとBプラン）の平均値（例：満足度）に差があるか検定。")
