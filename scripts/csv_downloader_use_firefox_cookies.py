"""マネーフォワードからcsvをダウンロードしてくるツール"""

import os
import sys
import time
from datetime import datetime
from io import StringIO

import browser_cookie3
import pandas as pd
import requests
from bs4 import BeautifulSoup
from dateutil.relativedelta import relativedelta


def assert_get(url: str, session: requests.Session) -> requests.Response:
    """GETリクエストが成功したかどうかを確認する"""
    response = session.get(url)
    assert response.status_code == 200, (
        "GETリクエストが失敗しました\nstatus code: " + str(response.status_code) + "\nurl: " + url
    )
    # お行儀をよくするために一秒まつ
    time.sleep(1)
    return response


def get_firefox_cookie(username: str, profile: str, wsl: bool = True) -> requests.Session | None:
    """マネーフォワードにwindows_firefoxからcookieを取得してログイン確認

    Args:
        username (str): User明
        profile (str): パスワード
        wsl (bool, optional): WSL上で実行するかどうか. Defaults to True.

    Returns:
        requests.Session: ログイン済みのセッション
    """

    sign_in_check = "https://moneyforward.com/"

    if wsl:
        cookie_file = os.path.join(
            os.sep,
            "mnt",
            "c",
            "Users",
            username,
            "AppData",
            "Roaming",
            "Mozilla",
            "Firefox",
            "Profiles",
            profile,
            "cookies.sqlite",
        )
    else:
        cookie_file = os.path.join(
            "C:\\",
            "Users",
            username,
            "AppData",
            "Roaming",
            "Mozilla",
            "Firefox",
            "Profiles",
            profile,
            "cookies.sqlite",
        )
    cookies = browser_cookie3.firefox(cookie_file=cookie_file)
    session = requests.Session()
    session.cookies.update(cookies)
    # 必要に応じてヘッダーを設定
    session.headers.update(
        {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:91.0) Gecko/20100101 Firefox/91.0",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }
    )

    # Money Forwardのページにリクエストを送信
    response = session.get(sign_in_check)

    # login の成功をcookieを見て確認する
    cookie_names = [cookie.name for cookie in response.cookies]
    if len(cookie_names) > 0:
        print("ログインに成功しました．")
        return session
    else:
        session.close()
        print("ログインに失敗しました．")
        return None


# download
def get_monthly_finances_csv(session: requests.Session, from_year: int, from_month: int, save_path: str) -> None:
    """マネーフォワードからCSVを取得する．

    Args:
        session (requests.Session): すでにログイン済みのセッション
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
    profile = "hogehoge.default-release"

    # ダウンロードする期間と場所
    from_year = 2022
    from_month = 5
    save_path = os.path.join("private", "downloaded_csv")

    # ログイン
    session = get_firefox_cookie(username=username, profile=profile, wsl=True)
    # ダウンロード
    get_monthly_finances_csv(session, from_year=from_year, from_month=from_month, save_path=save_path)
    # ログインsessionを閉じる
    session.close()
