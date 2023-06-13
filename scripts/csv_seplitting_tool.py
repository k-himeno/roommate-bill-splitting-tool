"""マネーフォワードから出力したのデータを分割するツール"""

import os
import sys
from datetime import date

import pandas as pd
from openpyxl import load_workbook


def get_bill_splitting_data(data, bill_splitting_data, bill_splitting_flag):
    """割り勘をするデータを取り出す．過去に入力したデータは除く．

    Args:
        data (DataFrame): 新しく記入したいデータ．
        bill_splitting_data (DataFrame): 過去に記入済みのデータ．
        bill_splitting_flag (list): リスト中の文字列がメモに含まれている場合，割り勘として扱う．

    Returns:
        data_frame: Dataから割り勘として扱うデータを切り出したもの．
    """

    # 割勘フラグが立っている物を抽出
    bill_splitting_idx = data[data["メモ"].str.contains(bill_splitting_flag, na=False)].index

    # 既に集約したデータは除く
    bill_splitting_idx = bill_splitting_idx[~bill_splitting_idx.isin(bill_splitting_data.index)]
    data_for_write = data.loc[bill_splitting_idx]
    if data_for_write.empty:
        sys.exit("全てのデータが記入済みです")
    return data_for_write


# 割り勘に必要なデータを切り出し，成形する．
def format_bill_splitting_data(data, user="U1"):
    """割り勘をするデータをフォーマットする

    Args:
        data (DataFrame): DataFrame型のデータ．
        user (str, optional): 割り勘の負担者．"U1" or "U2". Defaults to "U1".
    """

    columns_to_remove = ["計算対象", "保有金融機関", "振替"]
    data = data.drop(columns=columns_to_remove)

    # メモの中身を分割する．
    data = pd.concat([data, data["メモ"].str.split(bill_splitting_flag, expand=True)], axis=1).drop(columns="メモ")
    data.rename(columns={0: "メモ", 1: "割勘比率"}, inplace=True)

    # 割り勘比率を分割する．
    data = pd.concat([data, data["割勘比率"].str.split(":|;", expand=True)], axis=1).drop(columns="割勘比率")
    data.rename(columns={0: "U1 比率", 1: "U2 比率"}, inplace=True)

    data["同意Flag"] = ""
    data["修正Flag"] = ""
    data["清算済Flag"] = "FALSE"
    data["U1 負担"] = 0
    data["U2 負担"] = 0

    # 割勘比率を数値に変換する．
    data["U1 比率"] = pd.to_numeric(data["U1 比率"])
    data["U2 比率"] = pd.to_numeric(data["U2 比率"])
    data["金額（円）"] = pd.to_numeric(data["金額（円）"])
    data.rename(columns={"金額（円）": user + "金額 (円)"}, inplace=True)
    # 割勘金額を計算する．
    assert user + " 比率" in data.columns, "User名が間違っています"
    data_calc = data.filter(like="比率", axis=1)
    for user_i in data_calc.columns:
        print(user_i)
        if not user_i == user + " 比率":
            print(data[user_i] * data[user + "金額 (円)"])
            data[user_i.replace("比率", "負担")] = data[user_i] * data[user + "金額 (円)"] / data_calc.sum(axis=1)

    return data


# 割勘用のデータをExcel形式に成形し，保存する．
def save_bill_splitting_data(data, filename="bill_splitting.xlsx", user="U1"):
    """割り勘用のデータをExcel形式に成形し，保存する．

    Args:
        data (DataFrame): 割り勘用のデータ．
        bill_splitting_file (str): 保存先のExcelファイル名
        user (str, optional): 割り勘の負担者．"U1" or "U2". Defaults to "U1".
    """

    if not os.path.exists(os.path.dirname(filename)):
        os.makedirs(os.path.dirname(filename))

    if os.path.exists(filename):
        # ファイルが存在する場合は既存のワークブックを読み込む
        bill_splitting_data = pd.read_excel(filename, sheet_name=None, index_col=0)
        if user in bill_splitting_data.keys():
            bill_splitting_data = bill_splitting_data[user]
            with pd.ExcelWriter(filename, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
                bill_splitting_data.to_excel(
                    writer, sheet_name="_".join([user, "archive", date.today().strftime("%Y-%m-%d")])
                )
        else:
            bill_splitting_data = pd.DataFrame()
    else:
        bill_splitting_data = pd.DataFrame()

    # 割り勘をするデータを取り出す．過去に入力したデータは除く
    data_for_write = get_bill_splitting_data(data, bill_splitting_data, bill_splitting_flag)

    # データを成形する．
    data = format_bill_splitting_data(data_for_write, user=user)

    # データを結合する
    data = pd.concat([bill_splitting_data, data], axis=0)
    # dataを日付でソートする
    data = data.sort_values(by="日付", ascending=False)
    # DataFrameをExcelのSheet1に書き込む
    with pd.ExcelWriter(filename, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        data.to_excel(writer, sheet_name=user)


# os.chdir(os.path.abspath(__file__))
data_folder = os.path.join("data")

# 編集後のデータを保存するフォルダ
output_folder = os.path.join("output")

# 割り勘のためのフラグ
bill_splitting_flag = ["阿良々木割勘", "阿良々木割り勘"]
space = "\s*　*"
bill_splitting_flag = space + "|".join(bill_splitting_flag) + space

# マネーフォワードから入力されたファイルを読み込む．
# data_file_i = pd.read_csv(os.path.join(data_folder, "sample_data.csv"), encoding="shift-jis")

data_i = pd.read_csv(os.path.join(data_folder, "sample_data.csv"), encoding="shift-jis", index_col="ID")

data = data_i

save_bill_splitting_data(data, filename=os.path.join(output_folder, "_bill_splitting_sample.xlsx"), user="U1")
