from selenium import webdriver
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
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

def setup_logging():
    """ログ設定を初期化"""
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
    else:
        exe_dir = os.path.dirname(os.path.abspath(__file__))

    log_filename = f"automation_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
    log_path = os.path.join(exe_dir, log_filename)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_path), logging.StreamHandler()],
    )
    logging.info("処理を開始します。")
    return exe_dir

def load_config(exe_dir):
    """設定ファイルを読み込み"""
    try:
        with open(os.path.join(exe_dir, "config.json"), "r", encoding="utf-8") as file:
            config = json.load(file)
            logging.info("設定ファイルの読み込みが完了しました。")
            return config
    except Exception as e:
        logging.error(f"設定ファイルの読み込みに失敗しました: {e}")
        input("Enterを押して終了します...")
        sys.exit(1)

def setup_browser(config, exe_dir):
    """ブラウザを設定・起動"""
    proxy = config.get("proxy", "")
    
    # Edge オプション
    edge_options = EdgeOptions()
    edge_options.add_argument("--remote-allow-origins=*")
    edge_options.add_argument("--ignore-certificate-errors")
    edge_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    edge_options.add_argument("--disable-gpu")
    edge_options.add_argument("--no-sandbox")

    # プロキシ設定
    if proxy:
        os.environ["http_proxy"] = proxy
        os.environ["https_proxy"] = proxy
        os.environ["no_proxy"] = "127.0.0.1,localhost"
        edge_options.add_argument(f"--proxy-server={proxy}")
        logging.info(f"プロキシを設定しました: {proxy}")

        # 認証付きプロキシ等のための拡張（必要な環境のみ）
        root_dir = os.path.abspath(os.path.join(os.path.abspath(os.path.dirname(__file__)), "..", ".."))
        tmp_work_dir = os.path.abspath(os.path.join(root_dir, "tmpwork"))
        if not os.path.exists(tmp_work_dir):
            os.makedirs(tmp_work_dir)
        try:
            plugin_path = ProxyUtil.create_proxy_extentions(proxy, tmp_work_dir)
            edge_options.add_extension(plugin_path)
            logging.info("プロキシ拡張機能を追加しました。")
        except Exception as e:
            logging.warning(f"プロキシ拡張のロードに失敗しました（続行します）: {e}")

    # EdgeDriver 起動
    try:
        local_driver_path = os.path.join(exe_dir, "msedgedriver.exe")
        if os.path.exists(local_driver_path):
            logging.info(f"ローカルのEdgeDriverを使用します: {local_driver_path}")
            service = EdgeService(executable_path=local_driver_path)
            driver = webdriver.Edge(service=service, options=edge_options)
        else:
            logging.info("ローカルのmsedgedriver.exeが見つかりません。Selenium Manager で自動解決を試みます。")
            service = EdgeService()
            driver = webdriver.Edge(service=service, options=edge_options)
        logging.info("EdgeDriverの起動に成功しました。")
        return driver
    except Exception as e:
        logging.error(f"EdgeDriverの起動に失敗しました: {e}")
        logging.error("ネットワーク制限がある場合は、プロジェクト直下に Edge と同じメジャーバージョンの msedgedriver.exe を配置してください。")
        input("Enterを押して終了します...")
        sys.exit(1)

def login_to_website(driver, config):
    """ウェブサイトにログイン"""
    website_url = config["website_url"]
    username = config["login"]["username"]
    password = config["login"]["password"]
    
    try:
        driver.get(website_url)
        logging.info(f"ページを開きました: {website_url}")
        
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
        driver.find_element(By.ID, "password").send_keys(password)
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'ログインする')]"))
        )
        login_button.click()
        logging.info("ログイン処理を完了しました。")
    except Exception as e:
        logging.error(f"ログイン処理中にエラーが発生しました: {e}")
        driver.quit()
        input("Enterを押して終了します...")
        sys.exit(1)

def load_excel_data(excel_path):
    """Excelデータを読み込み"""
    try:
        # エクセルのsheet1からB1から対象月を取得する
        df_temp = pd.read_excel(excel_path, sheet_name="作業時間", header=None)
        if df_temp.empty:
            logging.info("Excelデータが空です。処理を終了します。")
            return None, None
        
        # 1行目の2列目（B1セル）に対象月がある想定
        target_month = df_temp.iloc[0, 1]  # .at[0, 1] から .iloc[0, 1] に修正
        logging.info(f"対象月を取得しました: {target_month}")
        
        # エクセルのsheet1の4行目から勤怠データの一覧を取得する
        df = pd.read_excel(excel_path, sheet_name="作業時間", skiprows=3)
        logging.info(f"勤怠データを読み込みました: {excel_path}")

        # 必要列のチェック
        required_cols = ["日付","曜日", "作業開始時刻","作業終了時刻", "作業時間", "作業内容", "備考", "確認", "プロジェクト", "アクティビティ", "説明", "タグ"]
        if not all(col in df.columns for col in required_cols):
            logging.error(f"Excelデータに必要な列がありません。必要列: {required_cols} / 実際: {list(df.columns)}")
            return None, None
        
        return target_month, df
    except Exception as e:
        logging.error(f"Excelデータの読み込みに失敗しました: {e}")
        return None, None

def format_time(time_value, default_time):
    """時刻を整形"""
    if pd.notna(time_value):
        if hasattr(time_value, 'strftime'):
            return time_value.strftime('%H:%M')
        else:
            time_str = str(time_value)
            if ':' in time_str and len(time_str) > 5:
                # "9:00:00" -> "09:00"
                parts = time_str.split(':')
                return f"{parts[0].zfill(2)}:{parts[1]}"
            else:
                return time_str.zfill(5)
    return default_time

def format_date(target_month, day):
    """日付を整形"""
    day_str = str(day).zfill(2)
    # 日が１～２０までの場合
    if day_str <= "20":
        return f"{str(target_month)[:4]}/{str(target_month)[4:6]}/{day_str}"
    else:
        if str(target_month)[4:6] == "1":
            next_month = f"{int(str(target_month)[:4])-1}12"
            return f"{next_month[:4]}/{next_month[4:6]}/{day_str}"
        else:
            return f"{str(target_month)[:4]}/{str(int(str(target_month)[4:6])-1).zfill(2)}/{day_str}"

def process_timesheet_entry(driver, row, target_month, df, index):
    """1行分のタイムシートエントリを処理"""
    try:
        # 作業時間が空ならスキップ
        if pd.isna(row["作業時間"]):
            logging.info(f"{index+1}行目の作業時間が空のためスキップします。")
            return True

        # 日付の処理
        full_date_str = format_date(target_month, row["日付"])
        
        # 時刻の整形
        start_time_str = format_time(row["作業開始時刻"], "09:00")
        end_time_str = format_time(row["作業終了時刻"], "17:00")

        # 新規作成
        BaseUtils.wait_and_click(driver, By.XPATH, "//a[contains(text(), '新規作成')]")
        time.sleep(2)
        logging.info(f"{index+1}行目のデータ入力を開始します。")

        # 各フィールドの入力
        BaseUtils.wait_and_send_keys(driver, By.ID, "timesheet_edit_form_begin_date", full_date_str)
        BaseUtils.wait_and_send_keys(driver, By.ID, "timesheet_edit_form_begin_time", start_time_str)
        BaseUtils.wait_and_send_keys(driver, By.ID, "timesheet_edit_form_duration", str(row["作業時間"]))

        # プロジェクト選択
        project_name = row["プロジェクト"]
        if project_name != "":
            BaseUtils.wait_and_select_value(driver, By.ID, "timesheet_edit_form_project-ts-control", project_name)
        else:
            BaseUtils.wait_and_select_value(driver, By.ID, "timesheet_edit_form_project-ts-control", "INES")

        # アクティビティ選択
        activity_name = row["アクティビティ"]
        BaseUtils.wait_and_select_value(driver, By.ID, "timesheet_edit_form_activity-ts-control", activity_name)

        # 作業内容（任意）
        if "作業内容" in df.columns and pd.notna(row["作業内容"]):
            BaseUtils.wait_and_send_keys(driver, By.ID, "timesheet_edit_form_description", str(row["作業内容"]))

        # タグ（任意）
        if "タグ" in df.columns and pd.notna(row["タグ"]):
            BaseUtils.wait_and_select_value(driver, By.ID, "timesheet_edit_form_tags-ts-control", str(row["タグ"]), escapeFlg=1)
        else:
            BaseUtils.wait_and_select_value(driver, By.ID, "timesheet_edit_form_tags-ts-control", "社員", escapeFlg=1)

        # 保存
        BaseUtils.wait_and_click(driver, By.XPATH, "//button[contains(text(), '保存する')]")
        logging.info(f"{index+1}行目のデータ入力が完了しました。")
        return True
    except Exception as e:
        logging.error(f"{index+1}行目でエラーが発生しました: {e}")
        return False

def main():
    """メイン処理"""
    # 初期化
    exe_dir = setup_logging()
    config = load_config(exe_dir)
    
    # ブラウザ起動・ログイン
    driver = setup_browser(config, exe_dir)
    login_to_website(driver, config)
    
    try:
        # Excelデータ読み込み
        target_month, df = load_excel_data(config["excel_path"])
        if target_month is None or df is None:
            driver.quit()
            input("Enterを押して終了します...")
            sys.exit(1)

        # 自動入力処理
        for index, row in df.iterrows():
            process_timesheet_entry(driver, row, target_month, df, index)

        logging.info("自動入力が完了しました。")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
