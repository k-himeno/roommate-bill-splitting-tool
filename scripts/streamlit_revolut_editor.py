# -*- coding: utf-8 -*-
"""
Streamlit版 家計簿アプリ：Revolut取引編集ツール
要件:
- 同一ディレクトリ内のCSVを統合してDataFrame化
- Completed Date が空欄の行は除外
- 最新のoutput_revolut_日付 CSV があれば読み込み・マージ
  - ID重複時は出力ファイル側を優先
  - pandas.concat で結合し、df_new側にない列をまとめて初期値埋め
  - 出力ファイル未存在時は新規行にまとめて初期列を追加
- EXCHANGE タイプの行を基にレート計算（有効数字6桁）
- (Amount + Fee) * rate で「金額（円）」列を追加
- 通貨ごとの Balance を再計算・書き換え（offset可）
- サイドバー：分類／中項目／割勘相手／計算対象 フィルター
- 列順指定 & st.data_editor で編集UI
- 未入力セルは赤背景でハイライト
- 編集結果をCSVとして保存
"""

import glob
import hashlib
import os
from datetime import datetime
from typing import Any, Dict, List

import numpy as np
import pandas as pd
import streamlit as st


@st.cache_data(show_spinner=False)
def load_revolut_csvs(data_dir: str, file_pattern: str) -> pd.DataFrame:
    """
    指定ディレクトリからRevolut取引CSVを読み込み、結合して返却する

    Args:
        data_dir (str): CSVファイルが格納されたディレクトリパス
        file_pattern (str): 読み込み対象ファイル名パターン（glob形式）

    Returns:
        pd.DataFrame: 結合後の取引データ。ファイルがなければ空のDataFrameを返す。
    """
    pattern = os.path.join(data_dir, file_pattern)
    file_list = glob.glob(pattern)
    dfs: List[pd.DataFrame] = []
    for fp in file_list:
        try:
            df = pd.read_csv(fp, parse_dates=["Started Date", "Completed Date"])
            dfs.append(df)
        except Exception as e:
            st.error(f"読み込みエラー: {fp} → {e}")
    if not dfs:
        return pd.DataFrame()

    # 結合
    df = pd.concat(dfs, ignore_index=True)
    df.dropna(subset=["Completed Date"], inplace=True)
    return df


def generate_unique_id(df: pd.DataFrame, rev_code: str) -> pd.DataFrame:
    """
    Started Date と Description を元に MD5 ハッシュを生成し、ID列を追加する

    Args:
        df (pd.DataFrame): ID付与対象のDataFrame
        rev_code (str): 3文字の金融機関コード

    Returns:
        pd.DataFrame: ID列が追加されたDataFrame
    """
    df = df.copy()

    def make_id(row: pd.Series) -> str:
        dt_str = row["Started Date"].strftime("%Y%m%d%H%M%S")
        desc = row["Description"]
        hash_value = hashlib.md5(desc.encode("utf-8")).hexdigest().upper()
        unique_id = "".join([dt_str, "_", hash_value[:6], rev_code])
        return unique_id

    df["ID"] = df.apply(make_id, axis=1)
    return df


@st.cache_data(show_spinner=False)
def load_and_merge_latest_output(
    data_dir: str, df_new: pd.DataFrame, output_prefix: str, defaults: Dict[str, Any]
) -> pd.DataFrame:
    """
    最新の出力ファイルと新規データをマージする

    Args:
        data_dir (str): 出力ファイルのあるディレクトリ
        df_new (pd.DataFrame): 最新取引データ
        output_prefix (str): 出力ファイル名のプレフィックス
        defaults (Dict[str, Any]): 新しい列に設定するデフォルト値マッピング

    Returns:
        pd.DataFrame: マージ後のDataFrame
    """
    pattern = os.path.join(data_dir, f"{output_prefix}*.csv")
    existing_files = glob.glob(pattern)
    if existing_files:
        latest_file = max(existing_files, key=os.path.getmtime)
        df_out = pd.read_csv(latest_file)
        # df_newにない列をまとめて追加
        missing_cols = {col: defaults[col] for col in defaults if col not in df_new.columns}
        if missing_cols:
            df_new = df_new.assign(**missing_cols)
        # ID重複分は出力側優先
        new_filtered = df_new[~df_new["ID"].isin(df_out["ID"])]
        merged = pd.concat([df_out, new_filtered], ignore_index=True, sort=False)
        return merged
    else:
        # 初回実行: defaultsを一括追加
        if defaults:
            df_new = df_new.assign(**defaults)
        return df_new


def recalculate_balances(df: pd.DataFrame, offset: float = 0.0) -> pd.DataFrame:
    """
    通貨ごとにStarted Date順でBalanceを再計算し書き換える。
    新しいBalance = 前Balance + Amount + offset

    Args:
        df (pd.DataFrame): 'Currency','Started Date','Amount' 列を含むデータ
        offset (float): 各行に加算するオフセット値 (デフォルト0.0)
    Returns:
        pd.DataFrame: 'Balance' 列を書き換えたDataFrame
    """
    df2 = df.sort_values("Started Date").reset_index(drop=True)

    # 各通貨グループで累積
    def _adjust(group: pd.DataFrame) -> pd.DataFrame:
        balances = []
        prev = 0.0
        for _, row in group.iterrows():
            new_bal = prev + row["Amount"] + offset
            balances.append(new_bal)
            prev = new_bal
        group["Balance"] = balances
        return group

    return df2.groupby("Currency", group_keys=False).apply(_adjust)


def calculate_exchange_rates(df: pd.DataFrame, base_currency: str = "JPY") -> pd.DataFrame:
    """
    EXCHANGE タイプの行を基に通貨ごとの円換算レートを計算し、exchange_rate 列を追加。

    Args:
        df (pd.DataFrame): 元データ。'Type','Started Date','Currency','Amount','Balance' 列必須。
        base_currency (str): 基軸通貨コード (デフォルト 'JPY')。

    Returns:
        pd.DataFrame: 'exchange_rate' 列追加済み DataFrame。
    """
    # 日付でソート
    df2 = df.sort_values("Started Date").reset_index(drop=True)
    # 基軸通貨は常に1.0
    last_rate: Dict[str, float] = {base_currency: 1.0}
    df2["exchange_rate"] = pd.NA

    # EXCHANGE 行だけ計算
    exch = df2[df2["Type"] == "EXCHANGE"]
    for t, group in exch.groupby("Started Date"):
        # 売り（Amount<0）の円換算合計
        sell_group = group[group["Amount"] < 0]
        total_sale_jpy = (
            sell_group["Amount"].abs() * sell_group["Currency"].map(lambda c: last_rate.get(c, 0.0))
        ).sum()

        # 買い（Amount>0）でレート更新
        buy_group = group[group["Amount"] > 0]
        for idx, row in buy_group.iterrows():
            cur = row["Currency"]
            amt = row["Amount"]
            bal = row["Balance"]
            old = last_rate.get(cur, 0.0)
            # 新レート計算式
            new_rate = round((old * (bal - amt) + total_sale_jpy) / bal, 4)
            last_rate[cur] = new_rate
            df2.at[idx, "exchange_rate"] = new_rate

        # 売りはレートをそのまま適用
        for idx, row in sell_group.iterrows():
            df2.at[idx, "exchange_rate"] = last_rate.get(row["Currency"], 0.0)

    # 通貨ごとに前回レートを埋め、JPY は常に 1.0 を保持
    df2["exchange_rate"] = (
        df2.groupby("Currency")["exchange_rate"]
        .ffill()
        .where(df2["Currency"] != base_currency)  # base_currency 以外は ffill
    )
    # base_currency の行は常に 1.0
    df2.loc[df2["Currency"] == base_currency, "exchange_rate"] = 1.0
    df2["金額（円）"] = df2["Amount"] + df2["Fee"] * df2["exchange_rate"]

    return df2


def save_output_csv(df: pd.DataFrame, data_dir: str, output_prefix: str) -> str:
    """
    DataFrameを本日日付のファイル名でCSV出力し、そのパスを返却

    Args:
        df (pd.DataFrame): 保存対象DataFrame
        data_dir (str): 保存先ディレクトリ
        output_prefix (str): 出力ファイル名のプレフィックス

    Returns:
        str: 書き出したCSVファイルのパス
    """
    date_str = datetime.now().strftime("%Y%m%d")
    filename = f"{output_prefix}{date_str}.csv"
    path = os.path.join(data_dir, filename)
    df.to_csv(path, index=False)
    return path


def main() -> None:
    st.set_page_config(page_title="Revolut 取引編集ツール", layout="wide")
    st.title("Revolut取引編集ツール")
    # ユーザー入力
    data_dir = st.text_input("CSV格納フォルダ", value=os.path.join(os.getcwd(), "private", "downloaded_csv"))
    file_pattern = st.text_input("Revolutファイルパターン", value="account-statement_*.csv")
    output_prefix = st.text_input("出力プレフィックス", value="output_revolut_")

    # 初期設定値
    defaults = {
        "大項目": "未分類",
        "中項目": "未分類",
        "保有金融機関": "Revolut",
        "計算対象": 1,
        "割勘対象": 0,
        "割勘相手": "",
        "自己負担比": "",
        "相手負担比": "",
    }
    rev_code = st.text_input("3文字コード", value="REV")

    # if "df_st" not in st.session_state:
    if st.button("処理開始"):
        with st.spinner("前処理を実行中..."):
            df_all = load_revolut_csvs(data_dir, file_pattern)
            df_id = generate_unique_id(df_all, rev_code)
            df = load_and_merge_latest_output(data_dir, df_id, output_prefix, defaults)
            df = recalculate_balances(df)
            df = calculate_exchange_rates(df)
            st.session_state.df_st = df
        st.success("前処理完了")
        st.rerun()  # 直後に rerun して edit_mode を初期 state に

    # ---- 以降：データがロード済みの場合のみ ----
    if "df_st" not in st.session_state:
        st.info("まず『処理開始』で CSV を読み込んでください。")
        return

    df_st = st.session_state.df_st

    desired_order = [
        "Type",
        "Started Date",
        "Description",
        "Amount",
        "Fee",
        "Currency",
        "大項目",
        "中項目",
        "保有金融機関",
        "計算対象",
        "割勘対象",
        "割勘相手",
        "自己負担比",
        "相手負担比",
        "金額（円）",
        "exchange_rate",
    ]
    # フィルター UI
    # st.sidebar.header("絞り込み")
    # for col in ["大項目", "中項目", "保有金融機関"]:
    #     opts = sorted(df_merged[col].unique())
    #     sel = st.sidebar.multiselect(f"{col} フィルター", opts, default=opts)
    #     df_merged = df_merged[df_merged[col].isin(sel)]

    # ---- 編集モード トグル ----
    edit_mode = st.toggle("編集モード", key="edit_mode", value=True)
    if edit_mode:
        with st.form("editor_form", clear_on_submit=False):
            edited_df = st.data_editor(
                df_st,
                use_container_width=True,
                column_order=[c for c in desired_order if c in df_st.columns],
                column_config={
                    "大項目": st.column_config.SelectboxColumn("大項目", options=["aa, bb", "cc", "dd"]),
                    "中項目": st.column_config.SelectboxColumn("中項目", options=["abc", "def"]),
                    "計算対象": st.column_config.CheckboxColumn("計算対象"),
                    "割勘対象": st.column_config.CheckboxColumn("割勘対象"),
                    "割勘相手": st.column_config.SelectboxColumn(
                        "割勘相手", options=["", "友人A", "友人B", "家族C"], default=""
                    ),
                    "自己負担比": st.column_config.TextColumn("自己負担比"),
                    "相手負担比": st.column_config.TextColumn("相手負担比"),
                },
                hide_index=True,
                key="editor",
            )
            if st.form_submit_button("保存して終了"):
                # 1) DataFrame を永続化
                st.session_state.df_st = edited_df
                csv_path = save_output_csv(edited_df, data_dir, output_prefix)
                st.success(f"編集内容を保存しました → {csv_path}")
                # 2) 閲覧モードへ遷移
                st.session_state.edit_mode = False
                st.rerun()
        # st.session_state.df_st = edited_df
    else:
        st.dataframe(df, use_container_width=True, column_order=[c for c in desired_order if c in df_st.columns])
        if st.button("CSV出力"):
            out_path = save_output_csv(df_st, data_dir, output_prefix)
            st.success(f"{out_path} に保存しました")


if __name__ == "__main__":
    main()
