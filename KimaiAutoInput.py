from selenium import webdriver
from selenium.webdriver.edge.service import Service
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys
import json
import os
from proxy_utils import ProxyUtil
from base_utils import BaseUtils
import time
import pandas as pd
import logging
from datetime import datetime

# exeと同じフォルダにログを保存
# 実行ファイルの場所を取得
if getattr(sys, "frozen", False):
    exe_dir = os.path.dirname(sys.executable)  # exeファイルがあるフォルダ
else:
    exe_dir = os.path.dirname(os.path.abspath(__file__))

# 日付付きのログファイルを作成
log_filename = f"automation_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
log_path = os.path.join(exe_dir, log_filename)

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_path),  # ファイルに出力
        logging.StreamHandler(),  # コンソールにも出力
    ],
)

logging.info("処理を開始します。")

# 設定ファイルを読み込む
try:
    with open(os.path.join(exe_dir, "config.json"), "r", encoding="utf-8") as file:
        config = json.load(file)
        logging.info("設定ファイルの読み込みが完了しました。")
except Exception as e:
    logging.error(f"設定ファイルの読み込みに失敗しました: {e}")
    input("Enterを押して終了します...")
    exit(1)

proxy = config["proxy"]  # プロキシ設定を取得
username = config["login"]["username"]
password = config["login"]["password"]
website_url = config["website_url"]
excel_path = config["excel_path"]

# Edgeのオプション設定
edge_options = webdriver.EdgeOptions()
edge_options.add_argument("--remote-allow-origins=*")
edge_options.add_argument("--ignore-certificate-errors")
edge_options.add_experimental_option("excludeSwitches", ["enable-logging"])
edge_options.add_argument("--disable-gpu")
edge_options.add_argument("--no-sandbox")

if proxy != "":
    os.environ["http_proxy"] = proxy
    os.environ["https_proxy"] = proxy
    os.environ["no_proxy"] = "127.0.0.1,localhost"
    edge_options.add_argument(f"--proxy-server={proxy}")
    logging.info(f"プロキシを設定しました: {proxy}")

    root_dir = os.path.abspath(
        os.path.join(os.path.abspath(os.path.dirname(__file__)), "..", "..")
    )
    tmp_work_dir = os.path.abspath(os.path.join(root_dir, "tmpwork"))
    if not os.path.exists(tmp_work_dir):
        os.makedirs(tmp_work_dir)
    plugin_path = ProxyUtil.create_proxy_extentions(proxy, tmp_work_dir)
    edge_options.add_extension(plugin_path)
    logging.info("プロキシ拡張機能を追加しました。")

# EdgeDriverの起動
try:
    service = Service(EdgeChromiumDriverManager().install())
    driver = webdriver.Edge(service=service, options=edge_options)
    logging.info("EdgeDriverの起動に成功しました。")
except Exception as e:
    logging.error(f"EdgeDriverの起動に失敗しました: {e}")
    input("Enterを押して終了します...")
    sys.exit(1)

# ページを開く
try:
    driver.get(website_url)
    logging.info(f"ページを開きました: {website_url}")
except Exception as e:
    logging.error(f"ページを開く際にエラーが発生しました: {e}")
    driver.quit()
    input("Enterを押して終了します...")
    sys.exit(1)

# ログイン処理
try:
    driver.find_element(By.ID, "username").send_keys(username)
    driver.find_element(By.ID, "password").send_keys(password)
    login_button = driver.find_element(
        By.XPATH, "//button[contains(text(), 'ログインする')]"
    )
    login_button.click()
    logging.info("ログイン処理を完了しました。")
except Exception as e:
    logging.error(f"ログイン処理中にエラーが発生しました: {e}")
    driver.quit()
    input("Enterを押して終了します...")
    sys.exit(1)

# Excelデータの読み込み
try:
    df = pd.read_excel(excel_path)
    logging.info(f"Excelデータを読み込みました: {excel_path}")
    # 空ファイルの場合、正常終了する
    if df.empty:
        logging.info("Excelデータが空です。処理を終了します。")
        driver.quit()
        input("Enterを押して終了します...")
        sys.exit(0)
    # 列「日付」「開始時間」「作業時間」「プロジェクト」「アクティビティ」がない場合、エラーにする
    if not all(
        col in df.columns
        for col in ["日付", "開始時間", "作業時間", "プロジェクト", "アクティビティ"]
    ):
        logging.error("Excelデータに必要な列がありません。")
        driver.quit()
        input("Enterを押して終了します...")
        sys.exit(1)
except Exception as e:
    logging.error(f"Excelデータの読み込みに失敗しました: {e}")
    driver.quit()
    input("Enterを押して終了します...")
    sys.exit(1)

# 自動入力処理
for index, row in df.iterrows():
    try:
        # 作業時間が空の場合、スキップ
        if pd.isna(row["作業時間"]):
            logging.info(f"{index+1}行目の作業時間が空のためスキップします。")
            continue

        # 新規作成ボタンをクリック
        BaseUtils.wait_and_click(driver, By.XPATH, "//a[contains(text(), '新規作成')]")
        time.sleep(2)
        logging.info(f"{index+1}行目のデータ入力を開始します。")
        # 日付
        BaseUtils.wait_and_send_keys(
            driver, By.ID, "timesheet_edit_form_begin_date", str(row["日付"])
        )

        # 開始時間
        BaseUtils.wait_and_send_keys(
            driver, By.ID, "timesheet_edit_form_begin_time", str(row["開始時間"])
        )

        # 作業時間
        BaseUtils.wait_and_send_keys(
            driver, By.ID, "timesheet_edit_form_duration", str(row["作業時間"])
        )
        # プロジェクト選択
        BaseUtils
        project_name = row["プロジェクト"]
        BaseUtils.wait_and_select_value(
            driver, By.ID, "timesheet_edit_form_project-ts-control", project_name
        )

        # アクティビティ
        activity_name = row["アクティビティ"]
        BaseUtils.wait_and_select_value(
            driver, By.ID, "timesheet_edit_form_activity-ts-control", activity_name
        )

        # 説明
        if pd.notna(row["説明"]):  # NaNでない場合
            description = row["説明"]
            BaseUtils.wait_and_send_keys(
                driver, By.ID, "timesheet_edit_form_description", description
            )
        # タグ
        if pd.notna(row["タグ"]):  # NaNでない場合
            tag = row["タグ"]
            BaseUtils.wait_and_select_value(
                driver, By.ID, "timesheet_edit_form_tags-ts-control", tag, escapeFlg=1
            )
        # 保存
        BaseUtils.wait_and_click(
            driver, By.XPATH, "//button[contains(text(), '保存する')]"
        )
        logging.info(f"{index+1}行目のデータ入力が完了しました。")

    except Exception as e:
        logging.error(f"{index+1}行目でエラーが発生しました: {e}")

logging.info("自動入力が完了しました。")
driver.quit()
