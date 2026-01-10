"""
マネーフォワードから出力したのデータを分割するツール
revolut にも対応
"""

from __future__ import annotations

import os
import sys
from datetime import date
from typing import Optional

import pandas as pd
from narwhals import Int64
from openpyxl import load_workbook

COLUMNS_TYPE_MAP = {
    "money_forward": {
        "ID": "string",
        "計算対象": bool,
        "日付": "string",  # YYYY/MM/DD
        "内容": "string",
        "金額（円）": float,
        "保有金融機関": "string",
        "大項目": "string",
        "中項目": "string",
        "メモ": "string",
        "振替": bool,
    },
    "revolut": {
        "source_id": "string",
        "Type": "string",
        "Started Date": str,  # YYYY-MM-DD HH:MM:SS
        "Completed Date": str,
        "Description": "string",
        "Amount": float,
        "Fee": float,
        "Currency": "string",
        "Balance": float,
        "exchange_rate": float,
        "MainCategory": "string",
        "SubCategory": "string",
        "memo": "string",
        "split_flag": "Int64",
        "split_partner_name": "string",
        "my_rate": float,
        "partner_rate": float,
        "bank_name": "string",
    },
    "bill_splitting": {
        "ID": str,
        "日付": str,
        "内容": "string",
        "金額（円）": float,
        "大項目": "string",
        "中項目": "string",
        "メモ": "string",
        "my_rate": float,
        "partner_rate": float,
        "reject": "Int64",
        "fixed": "Int64",
        "清算済": "Int64",
        "my_share": float,
        "partner_share": float,
    },
}


def read_csv_from_money_forward(data_folder: str, encoding: str = "utf-8") -> pd.DataFrame:
    """マネーフォワードから出力されたcsvファイルを読み込む．

    Args:
        data_folder (str): フォルダ名
        encoding (str, optional): エンコード.文字化けするときはshift-jisに． Defaults to "utf-8".

    Returns:
        pd.DataFrame: 読み込まれたデータ
    """
    first = True
    for data_file_i in os.listdir(data_folder):
        if data_file_i.endswith(".csv"):
            data_i = pd.read_csv(
                os.path.join(data_folder, data_file_i),
                encoding=encoding,
                index_col="ID",
                dtype=COLUMNS_TYPE_MAP["money_forward"],
            )
        if first:
            data = data_i
            first = False
        else:
            data = pd.concat([data, data_i], axis=0)
    data["日付"] = pd.to_datetime(data["日付"], format="%Y/%m/%d")
    data["日付"] = data["日付"].dt.strftime("%Y-%m-%d 00:00:00")

    # メモ列をNFKC正規化
    data["メモ"] = data["メモ"].fillna("").astype(str).str.normalize("NFKC").str.strip()
    return data


def money_forward_bill_split(data: pd.DataFrame, splitting_pattern: str) -> pd.DataFrame:
    """割り勘をするデータを取り出す

    Args:
        data (DataFrame): 新しく記入したいデータ．
        splitting_master (DataFrame): 割り勘マスタ．

    Returns:
        data_frame: Dataから割り勘として扱うデータを切り出したもの．
    """
    regex = None

    bill_splitting_idx = data[data["メモ"].str.contains(splitting_pattern, na=False)].index
    data_splitting_df = data.loc[bill_splitting_idx, ["日付", "内容", "金額（円）", "大項目", "中項目", "メモ"]].copy()

    # メモ列の分割
    memo_tmp = data_splitting_df["メモ"].str.split(splitting_pattern, expand=True, regex=regex, n=1)
    rate_tmp = memo_tmp[1].str.split(":|;", expand=True, n=1)
    data_splitting_df["メモ"] = memo_tmp[0]
    data_splitting_df["my_rate"] = rate_tmp[0]
    data_splitting_df["partner_rate"] = rate_tmp[1]

    # 正規化．
    data_splitting_df["my_rate"] = pd.to_numeric(data_splitting_df["my_rate"])
    data_splitting_df["partner_rate"] = pd.to_numeric(data_splitting_df["partner_rate"])
    data_splitting_df["メモ"] = data_splitting_df["メモ"].str.strip()

    return data_splitting_df


def revolut_bill_split(
    csv_path: str,
    split_partner_name: Optional[str] = None,
) -> pd.DataFrame:
    """正規化済み revolut.csv から割勘対象のみ抽出し、既存形式に合わせる。

    入力想定列:
      source_id, transaction_date, description, amount, memo, my_rate, partner_rate,
      split_flag, split_partner_name
    出力:
      index=source_id, 日付, 内容, 金額（円）, メモ, U1 比率, U2 比率, 銀行名,
    Args:
        csv_path (str): 正規化済み revolut.csv ファイルパス
        split_partner_name (Optional[str], optional): 指定した相手との割勘のみ抽出. Defaults to None.
        user (str, optional): 負担者. Defaults to "U1".
    """
    df = pd.read_csv(csv_path, dtype=COLUMNS_TYPE_MAP["revolut"])
    # 日付形式は統一済み

    # 割勘対象のみ
    df = df[df["split_flag"] == 1]
    if split_partner_name is not None:
        # 指定した相手との割勘のみ．
        # 完全一致
        df = df[df["split_partner_name"] == split_partner_name]
    # 既存列名に合わせる
    gross = df["Amount"] + df["Fee"]
    amt_jpy = gross * df["exchange_rate"]
    out = pd.DataFrame(
        {
            "ID": df["source_id"],
            "日付": df["Started Date"],
            "内容": df["Description"],
            "大項目": df["MainCategory"],
            "中項目": df["SubCategory"],
            "金額（円）": amt_jpy,
            "メモ": df["memo"],
            "my_rate": df["my_rate"],
            "partner_rate": df["partner_rate"],
        },
    )
    out = out.set_index("ID")

    return out


# 割勘比率に基づき，按分計算を行う．
def calculate_bill_splitting(data_splitting_df: pd.DataFrame) -> pd.DataFrame:
    """割勘比率に基づき，按分計算を行う．
    Args:
        data_splitting_df (DataFrame): 割勘対象データ
    Returns:
        DataFrame: 按分計算済みデータ
    """
    den = data_splitting_df["my_rate"] + data_splitting_df["partner_rate"]
    data_splitting_df["my_share"] = data_splitting_df["金額（円）"] * data_splitting_df["my_rate"] / den
    data_splitting_df["partner_share"] = data_splitting_df["金額（円）"] * data_splitting_df["partner_rate"] / den

    data_splitting_df["reject"] = ""
    data_splitting_df["fixed"] = ""
    data_splitting_df["清算済"] = 0

    return data_splitting_df


def save_bill_splitting_data(split_df: pd.DataFrame, xlsx_path: str, sheet_name: str, sort_by: str = "日付") -> None:
    """割勘DFを既存シートに追記して保存。
    既存シートが無ければ新規作成。
    あればアーカイブシートに退避。
    indexをID列として扱う。
    Args:
        split_df (DataFrame): 追加するデータ
        xlsx_path (str): 保存先Excelファイルパス
        sheet_name (str): シート名
        sort_by (str, optional): ソートする列名. Defaults to "日付".
    """
    xlsx_path = str(xlsx_path)
    if os.path.exists(xlsx_path):
        existed = pd.read_excel(xlsx_path, sheet_name=sheet_name, index_col=0)
        existed["日付"] = pd.to_datetime(existed["日付"], errors="coerce").dt.strftime("%Y-%m-%d 00:00:00")
    else:
        existed = pd.DataFrame()

    new_only = split_df.loc[~split_df.index.isin(existed.index)].copy()

    if new_only.empty:
        print("追加するデータはありません。")
        return

    combined = pd.concat([existed, new_only], axis=0)

    if sort_by in combined.columns:
        combined = combined.sort_values(sort_by, ascending=False)

    # 既存をアーカイブ退避
    if not existed.empty:
        with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a", if_sheet_exists="new") as w:
            existed.to_excel(w, sheet_name=f"{sheet_name}_archive_{date.today():%Y-%m-%d}")

    combined.index.name = "ID"

    # 置換保存
    with pd.ExcelWriter(xlsx_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        combined.to_excel(w, sheet_name=sheet_name, index=True, na_rep="")


# ======== 追加：revolut用ユーティリティ =======================


if __name__ == "__main__":
    # os.chdir(os.path.abspath(__file__))
    money_forward_folder = os.path.join("data")

    # 編集後のデータを保存するフォルダ
    output_folder = os.path.join("output")

    # 割り勘のためのフラグ
    bill_splitting_flag = ["阿良々木割勘", "阿良々木割り勘"]
    space = "\s*"

    bill_splitting_flag = ["".join([space, flag, space]) for flag in bill_splitting_flag]
    bill_splitting_str = "|".join(bill_splitting_flag)

    # マネーフォワードから入力されたファイルを読み込む．
    # ブラウザを使い手動でダウンロードしてきた場合は，エンコードをshift-jisにする．
    data = read_csv_from_money_forward(money_forward_folder, encoding="shift-jis")

    data_splitting_MF = money_forward_bill_split(data, bill_splitting_str)

    data = revolut_bill_split(
        "revolut_normalized.csv",
        split_partner_name="阿良々木",
    )
    data_splitting_Revolut = calculate_bill_splitting(data)

    data = pd.concat([data_splitting_MF, data_splitting_Revolut], axis=0)

    save_bill_splitting_data(
        data,
        xlsx_path=os.path.join(output_folder, "_bill_splitting_sample.xlsx"),
        sheet_name="U1",
    )
