"""マネーフォワードから出力したのデータを分割するツール"""

import os
import sys
import time
from datetime import datetime
from io import StringIO

import pandas as pd
import requests
from bs4 import BeautifulSoup
from dateutil.relativedelta import relativedelta


def assert_get(url, session):
    """GETリクエストが成功したかどうかを確認する"""
    response = session.get(url)
    assert response.status_code == 200, "GETリクエストが失敗しました\nstatus code: " + str(response.status_code) + "\nurl: " + url
    # お行儀をよくするために一秒まつ
    time.sleep(1)
    return response


def start_mf_session(username, password):
    """マネーフォワードにログインしたsessionを張る

    Args:
        username (str): ID
        password (str): パスワード
    """

    login_url = "https://moneyforward.com/login"
    sign_in_url = "https://moneyforward.com/sign_in"
    mail_login_url = "https://id.moneyforward.com/sign_in/email?"

    # セッションのインスタンスを作成する。
    session = requests.Session()

    # login 画面
    login_response = assert_get(login_url, session)
    login_soup = BeautifulSoup(login_response.text)

    # sign_in 画面 の script タグの中にあるパスを取得する。メールでログインするため．
    sign_in_response = assert_get(sign_in_url, session)
    sign_in_soup = BeautifulSoup(sign_in_response.text)
    mail_login_path = sign_in_soup.find("script")

    # query を取得する
    mail_login_path = str(mail_login_path).split("\n")[2].split(";")
    flag = False
    for mail_login_path_i in mail_login_path:
        if "gon.authorizationParamsQueryString" in mail_login_path_i:
            mail_login_query = mail_login_path_i
            # クエリの部分だけを取得する
            mail_login_query = mail_login_query.split('="')[1][:-1].replace("\\u0026", "&")
            flag = True
        if "gon.authorizationParams=" in mail_login_path_i:  # 辞書型の文字列を取得する
            mail_login_params = mail_login_path_i.split("=")[1]
    assert flag, "sign in ページのソースに gon.authorizationParamsQueryString がありません．"

    # query を整形する
    # トークンを取得するために，ログインページを取得する
    mail_login_url_response = assert_get(mail_login_url + mail_login_query, session)
    mail_login_url_soup = BeautifulSoup(mail_login_url_response.text)
    authenticity_token = mail_login_url_soup.find("meta", {"name": "csrf-token"})["content"]

    # post用にクエリを成形する．
    mail_login_params = eval(mail_login_params)
    mail_login_params |= {
        "mfid_user[email]": username,
        "mfid_user[password]": password,
        "authenticity_token": authenticity_token,
    }

    # loginをする
    top_page_response = session.post(
        url="https://id.moneyforward.com/sign_in",
        data=mail_login_params,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    assert top_page_response.status_code == 200, "POSTリクエストが失敗しました．\nstatus code: " + str(top_page_response.status_code)
    top_page_soup = BeautifulSoup(top_page_response.text)

    return session


# download
def get_monthly_finances_csv(session, from_year, from_month, save_path: str):
    """マネーフォワードからCSVを取得する．

    Args:
        session (_type_): すでにログイン済みのセッション
        from_year (int): 開始年
        from_month (int): 開始月
        save_path (str): 保存するフォルダ名
    """

    if not os.path.exists(save_path):
        os.makedirs(save_path)

    url = "https://moneyforward.com/cf/csv?"
    data_date = datetime(from_year, from_month, 1)
    today = datetime.today()

    while data_date <= today:
        url_with_date = (
            url
            + "from="
            + data_date.strftime("%Y/%m/%d")
            + "&month="
            + str(data_date.month)
            + "&year="
            + str(data_date.year)
        )
        response = assert_get(url_with_date, session)

        # レスポンスの内容を取得し、pandasのread_csvでデータを読み込みます
        content = response.content.decode("cp932")  # 必要に応じて文字エンコーディングを変更してください
        df = pd.read_csv(StringIO(content))  # 文字列をファイルとして扱い読み込む
        df.to_csv(os.path.join(save_path, data_date.strftime("%Y-%m") + ".csv"), encoding="utf_8_sig", index=False)

        data_date += relativedelta(months=1)


if __name__ == "__main__":
    # 適宜書き換え
    # login情報
    username = "hogehoge"
    password = "1234"

    # ダウンロードする期間と場所
    from_year = 2022
    from_month = 5
    save_path = os.path.join("private", "downloaded_csv")

    # ログイン
    session = start_mf_session(username=username, password=password)
    # ダウンロード
    get_monthly_finances_csv(session, from_year=from_year, from_month=from_month, save_path=save_path)
    # ログインsessionを閉じる
    session.close()
