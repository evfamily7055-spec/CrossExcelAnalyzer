import streamlit as st
import pandas as pd
import io

# Set page layout to wide
st.set_page_config(layout="wide")
st.title("Excelピボットテーブルジェネレーター 📊")

st.info("Excelファイルをアップロードすると、列名を自動抽出し、ピボットテーブルを作成します。")

# 1. File Uploader
uploaded_file = st.file_uploader("Excelファイル (.xlsx) をアップロードしてください", type=["xlsx"])

if uploaded_file:
    st.header("1. データプレビュー")
    try:
        # Read the Excel file
        # Use io.BytesIO to handle the uploaded file object in memory
        bytes_data = uploaded_file.getvalue()
        df = pd.read_excel(io.BytesIO(bytes_data))
        
        st.dataframe(df.head(10))
        
        st.header("2. ピボットテーブル設定")
        
        # Get column names for selectors
        columns = df.columns.tolist()
        
        # --- Sidebar for options ---
        st.sidebar.header("ピボット設定")
        
        # 2. Select Columns for Rows (Index)
        index_cols = st.sidebar.multiselect(
            "行 (Rows) に設定する項目を選択", 
            options=columns,
            help="ピボットテーブルの行（インデックス）になる列。"
        )
        
        # 3. Select Columns for Columns
        column_cols = st.sidebar.multiselect(
            "列 (Columns) に設定する項目を選択", 
            options=columns,
            help="ピボットテーブルの列になる列。（オプション）"
        )
        
        # 4. Select Column for Values
        value_col = st.sidebar.selectbox(
            "値 (Values) に設定する項目を選択", 
            options=columns, 
            index=None,
            placeholder="集計する列を選択...",
            help="集計対象となる数値データが含まれる列。"
        )
        
        # 5. Select Aggregation Function
        agg_func_options = {
            "合計 (Sum)": "sum",
            "平均 (Mean)": "mean",
            "カウント (Count)": "count",
            "中央値 (Median)": "median",
            "最小値 (Min)": "min",
            "最大値 (Max)": "max"
        }
        agg_func_label = st.sidebar.selectbox(
            "集計方法 (Aggregation)", 
            options=agg_func_options.keys(),
            help="値をどのように集計するか選択します。"
        )
        
        # Get the actual pandas function name
        agg_func = agg_func_options.get(agg_func_label, "sum")

        # --- End of Sidebar ---

        # 6. Generate Pivot Table
        if index_cols and value_col:
            st.header("3. ピボットテーブル実行結果")
            
            with st.spinner("ピボットテーブルを作成中..."):
                try:
                    # Create pivot table
                    pivot_df = pd.pivot_table(
                        df,
                        index=index_cols,
                        columns=column_cols if column_cols else None, # Handle empty column selection
                        values=value_col,
                        aggfunc=agg_func,
                        fill_value=0 # Fill NaN with 0 for cleaner output
                    )
                    
                    st.dataframe(pivot_df, use_container_width=True)
                    
                    # --- Add download button ---
                    
                    # Cache the conversion function
                    @st.cache_data
                    def convert_df_to_csv(df_to_convert):
                        # Use utf-8-sig to ensure correct encoding for CSV, especially for Japanese characters
                        return df_to_convert.to_csv(index=True).encode('utf-8-sig')

                    csv_data = convert_df_to_csv(pivot_df)
                    
                    st.download_button(
                        label="結果をCSVとしてダウンロード",
                        data=csv_data,
                        file_name="pivot_table_result.csv",
                        mime="text/csv",
                    )

                except Exception as e:
                    st.error(f"ピボットテーブルの作成に失敗しました: {e}")
                    st.error("（ヒント: '値' には数値データ列を、'集計方法' には '合計' や '平均' を選んでいますか？ テキストデータには 'カウント' をお試しください。）")

        else:
            # Guide user to make selections
            st.warning("サイドバーで「行」と「値」の項目を最低1つずつ選択してください。")

    except Exception as e:
        st.error(f"Excelファイルの読み込みに失敗しました: {e}")

else:
    st.info("👆 上のボタンからExcelファイルをアップロードして開始します。")
