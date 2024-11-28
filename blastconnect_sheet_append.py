import streamlit as st
import pandas as pd
from io import BytesIO
import base64
from datetime import datetime



bat_order_dic = {
    # 一塁側チームオーダー
    "2302 Player": "h_1",
    "2303 Player": "h_2",
    "2304 Player": "h_3",
    "2305 Player": "h_4",
    "2306 Player": "h_5",
    "2307 Player": "h_6",
    "2308 Player": "h_7",
    "2309 Player": "h_8",
    "2310 Player": "h_9",
    # 三塁側チームオーダー
    "2322 Player": "a_1",
    "2323 Player": "a_2",
    "2324 Player": "a_3",
    "2325 Player": "a_4",
    "2326 Player": "a_5",
    "2327 Player": "a_6",
    "2328 Player": "a_7",
    "2329 Player": "a_8",
    "2310 Player": "a_9",
}


def make_unique_columns(columns):
    """
    重複する列名に連番を付けて一意にする関数
    """
    seen = {}
    unique_columns = []

    for item in columns:
        # None や nan の場合は空文字列として処理
        item = str(item) if pd.notna(item) else ""
        if item in seen:
            seen[item] += 1
            unique_columns.append(f"{item}_{seen[item]}")
        else:
            seen[item] = 0
            unique_columns.append(item)

    return unique_columns


def process_excel_file(uploaded_file):
    """
    Excelファイルの各シートからデータを抽出し、
    最初のシートの9行目をヘッダーとして使用し、
    その他のシートの10行目以降のデータを結合する関数
    """
    try:
        # Excelファイルを読み込む
        excel_file = pd.ExcelFile(uploaded_file)

        # 全シートのデータを格納するリスト
        all_data = []

        # ヘッダーとなる列名を格納する変数
        column_headers = None

        # 処理状況を表示するプログレスバー
        progress_bar = st.progress(0)

        # 処理対象のシート名を取得（「Player」を含むシートのみ）
        sheet_names = [
            name
            for name in excel_file.sheet_names
            # if "Player" in name and name != "チームレポート"
            if name != "チームレポート"
        ]

        if not sheet_names:
            st.error("処理対象のシートが見つかりません。")
            return None

        # 各シートを処理
        for i, sheet_name in enumerate(sheet_names):
            try:
                # 進捗状況を更新
                progress = (i + 1) / len(sheet_names)
                progress_bar.progress(progress)

                # シート処理状況を表示
                st.write(f"処理中のシート: {sheet_name}")

                # シートを読み込む（ヘッダーなしで読み込む）
                df = pd.read_excel(
                    uploaded_file,
                    sheet_name=sheet_name,
                    header=None,
                    dtype={0: str},  # 最初の列を文字列として読み込む
                )

                # 空の行を削除
                df = df.dropna(how="all")
                # df = df.dropna(subset=4)

                if i == 0:  # 最初のシート
                    if len(df) > 8:
                        # 9行目（インデックス8）を列名として取得し、重複を処理
                        raw_headers = df.iloc[7].fillna("").tolist()
                        column_headers = make_unique_columns(raw_headers)

                        # 列名の重複状況を表示
                        duplicates = [
                            x
                            for x in raw_headers
                            if raw_headers.count(x) > 1 and x != ""
                        ]
                        if duplicates:
                            st.warning(
                                f"重複する列名が検出されました: {set(duplicates)}"
                            )
                            st.write("重複する列名には連番が付与されます")

                        # 10行目以降のデータを抽出
                        if len(df) > 8:
                            data_df = df.iloc[:].copy()
                            data_df.columns = column_headers

                            # 日付列（最初の列）の処理
                            try:
                                df.iloc[0:].copy()
                            except Exception as date_error:
                                st.warning(
                                    f"日付の変換中にエラーが発生しました: {str(date_error)}"
                                )

                            data_df["元のシート名"] = sheet_name
                            all_data.append(data_df)
                    else:
                        st.error(
                            f"最初のシート '{sheet_name}' に9行以上のデータがありません。"
                        )
                        return None
                else:  # 2番目以降のシート
                    if len(df) > 8:
                        # 10行目以降のデータを抽出
                        data_df = df.iloc[8:].copy()

                        # 列数を調整
                        if len(data_df.columns) >= len(column_headers):
                            data_df = data_df.iloc[:, : len(column_headers)]
                        else:
                            # 足りない列を追加
                            for _ in range(len(column_headers) - len(data_df.columns)):
                                data_df[f"空列_{_}"] = pd.NA

                        data_df.columns = column_headers

                        data_df["元のシート名"] = sheet_name
                        data_df["bat_order"] = data_df["元のシート名"].replace(
                            bat_order_dic
                        )
                        # 試合以外を除外
                        data_df = data_df[data_df["スイング条件"]=="In Game"]
                        all_data.append(data_df)

            except Exception as sheet_error:
                st.warning(
                    f"シート '{sheet_name}' の処理中にエラーが発生しました: {str(sheet_error)}"
                )
                continue

        # プログレスバーを完了状態に
        progress_bar.progress(1.0)

        # データが存在する場合のみ結合処理を行う
        if all_data:
            # すべてのデータを結合
            merged_df = pd.concat(all_data, ignore_index=True)

            # 日付でソート
            try:
                merged_df = merged_df.sort_values(by=merged_df.columns[0])

                # 無効な日付（NaT）を持つ行を確認
                invalid_dates = merged_df[merged_df.iloc[:, 0].isna()]
                if not invalid_dates.empty:
                    st.warning(f"無効な日付を含む行が {len(invalid_dates)} 件あります")
            except Exception as sort_error:
                st.warning(f"日付でのソート中にエラーが発生しました: {str(sort_error)}")

            return merged_df
        else:
            st.warning("処理対象のデータが見つかりませんでした。")
            return None

    except Exception as e:
        st.error(f"エラーが発生しました: {str(e)}")
        return None


def to_excel_download_link(df, filename="processed_data.xlsx"):
    """
    DataFrameをExcelファイルとしてダウンロードするリンクを生成する関数
    """
    output = BytesIO()
    with pd.ExcelWriter(
        output, engine="openpyxl", datetime_format="YYYY-MM-DD"
    ) as writer:
        df.to_excel(writer, sheet_name="まとめ", index=False)

    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    href = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
    return href


# Streamlitアプリのメイン部分
def main():
    st.title("Blast Connectダウンロードデータまとめツール")
    st.write("複数シートのデータを1シートにまとめるツールです")
    st.write("※「チームレポート」シートはスキップされます")
    st.write("※最初のシートの9行目が列タイトルとして使用されます")
    st.write("※最初の列は日付データとして処理されます")

    # ファイルアップロード
    uploaded_file = st.file_uploader(
        "Excelファイルを選択してください", type=["xlsx", "xls"]
    )

    if uploaded_file is not None:
        # 実行ボタン
        if st.button("データを処理"):
            st.write("処理を開始します...")

            # データ処理
            processed_df = process_excel_file(uploaded_file)

            if processed_df is not None:
                # 処理結果のプレビューを表示
                st.write("処理結果のプレビュー（先頭5行）:")
                st.dataframe(processed_df.head())

                # データの概要を表示
                st.write("データの概要:")
                st.write(f"- 合計行数: {len(processed_df)}")
                st.write(f"- 列数: {len(processed_df.columns)}")
                st.write(f"- シート数: {len(processed_df['元のシート名'].unique())}")
                # st.write(
                #     f"- 日付範囲: {processed_df.iloc[:, 0].min()} から {processed_df.iloc[:, 0].max()}"
                # )

                # ダウンロードリンクを生成
                excel_link = to_excel_download_link(processed_df)

                # ダウンロードボタンを表示
                st.markdown(
                    f'<a href="{excel_link}" download="processed_data.xlsx">'
                    f'<button style="background-color:#4CAF50;color:white;padding:10px;'
                    f'border:none;border-radius:5px;cursor:pointer;">'
                    f"処理結果をダウンロード</button></a>",
                    unsafe_allow_html=True,
                )


if __name__ == "__main__":
    main()
