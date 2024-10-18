import flet as ft
import random
import time
import threading
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import gspread
from oauth2client.service_account import ServiceAccountCredentials

print("vs code で編集")
def get_item_ids_from_spreadsheet(credentials_path, spreadsheet_name, sheet_name):
    # Google APIの認証情報の設定
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_path, scope)
    client = gspread.authorize(creds)

    # Googleスプレッドシートの取得
    sheet = client.open(spreadsheet_name).worksheet(sheet_name)

    # スプレッドシートのデータ範囲を取得
    data = sheet.get_all_values()

    # item_idリストを作成（ヘッダーを除く）
    minne_selling_id_list = []
    for row in data[1:]:  # 最初の行（ヘッダー）を除く
        if len(row) > 1 and row[1]:  # 2列目にitem_idがあると仮定
            minne_selling_id_list.append(row[1])

    return minne_selling_id_list

# スプレッドシートから最新の全商品IDを取得する
credentials_path = "feisty-coast-438212-j4-dd46e5c40b12.json"  # 認証情報のパスを指定
spreadsheet_name = "ミンネ出品中リスト"  # スプレッドシートの名前を指定
sheet_name = "minne_sellimg_id_list"  # タブの名前を指定
minne_selling_id_list = get_item_ids_from_spreadsheet(credentials_path, spreadsheet_name, sheet_name)


# ログイン処理の関数
def minne_login(driver, status_text):
    status_text.value += "ログイン中...\n"
    status_text.update()
    driver.get("https://minne.com/signin")
    time.sleep(3)

    try:
        # メールアドレスとパスワードの入力
        elem_username = driver.find_element(By.ID, 'user_email')
        elem_username.send_keys('info@monster-girl.co.jp')
        elem_password = driver.find_element(By.ID, 'user_password')
        elem_password.send_keys('290102mi')

        time.sleep(1)

        # ログインボタンを探す
        elem_login_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "div.c-form__action input[type='submit']"))
        )
        # ログインボタンをクリックする
        elem_login_btn.click()

        time.sleep(3)

        status_text.value += "ログイン成功\n"
        status_text.update()
        print("Login successful")
        return True
    except Exception as ex:
        status_text.value += f"ログイン失敗: {str(ex)}\n"
        status_text.update()
        print(f"Login failed: {str(ex)}")
        return False

# 非公開にする関数
def minne_relist_off(driver, item_id):
    try:
        driver.get("https://minne.com/account/products/" + item_id)
        time.sleep(3)  # ページが完全に読み込まれるのを待つ

        # 商品タイトルの取得
        try:
            title_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "product_product_name"))
            )
            title = title_input.get_attribute("value")
            if not title:
                raise ValueError("タイトルが取得できませんでした。")
        except Exception as ex:
            raise RuntimeError(f"タイトル取得中にエラーが発生しました: {str(ex)}")

        # 公開停止ボタンをクリック
        dsp_btn_off = driver.find_element(By.XPATH, '//*[@id="product_disp_flg_false_label"]')
        dsp_btn_off.click()
        time.sleep(3)

        try:
            elem_login_btn = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[@class="p-section__action"]/input[@type="submit"]'))
            )
            elem_login_btn.click()
            time.sleep(3)
        except Exception as ex:
            raise RuntimeError(f"非公開ボタンのクリックに失敗しました: {str(ex)}")

        return title
    except Exception as ex:
        error_message = f"エラーが発生しました - {str(ex)}\n" + traceback.format_exc()
        print(error_message)
        return None

# 公開に戻す関数
def minne_relist_on(driver, item_id):
    try:
        driver.get("https://minne.com/account/products/" + item_id)
        time.sleep(3)

        try:
            dsp_btn_on = driver.find_element(By.XPATH, '//*[@id="product_disp_flg_true_label"]')
            dsp_btn_on.click()
            time.sleep(3)

            elem_login_btn = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[@class="p-section__action"]/input[@type="submit"]'))
            )
            elem_login_btn.click()
            time.sleep(3)
        except Exception as ex:
            raise RuntimeError(f"公開ボタンのクリックに失敗しました: {str(ex)}")
    except Exception as ex:
        error_message = f"エラーが発生しました - {str(ex)}\n" + traceback.format_exc()
        print(error_message)
        return None

import traceback

def execute_task(status_text, progress_bar, product_count_value, update_interval_value, update_count_value):
    try:
        count = int(product_count_value)
        update_interval = int(update_interval_value)
        update_count = int(update_count_value)

        for i in range(update_count):
            # WebDriverの設定（例：Chrome）
            options = Options()
            options.add_argument('--headless=old')
            options.add_argument('--window-size=1000,1000')
            service = Service()  # 自動更新　https://javeo.jp/no-more-webdriver_manager/
            driver = webdriver.Chrome(options=options, service=service)

            try:

                # ログイン処理
                if not minne_login(driver, status_text):
                    status_text.value += "ログイン失敗\n"
                    status_text.update()
                    print("Login failed.")
                    driver.quit()
                    return

                selected_ids = random.sample(minne_selling_id_list, count)
                status_text.value += f"更新 {i + 1} / {update_count} 開始\n"
                status_text.update()

                for item_id in selected_ids:
                    try:
                        title = minne_relist_off(driver, item_id)
                        if title:
                            status_text.value += f"{title} の公開停止\n"
                            print(f"{title} の公開停止")
                            status_text.update()

                            minne_relist_on(driver, item_id)
                            status_text.value += f"{title} の再公開\n"
                            print(f"{title} の再公開")
                            status_text.update()
                        else:
                            raise RuntimeError("公開停止処理に失敗しました。")
                    except Exception as item_ex:
                        error_message = f"商品ID {item_id} の処理中にエラーが発生しました: {str(item_ex)}\n" + traceback.format_exc()
                        status_text.value += error_message
                        print(error_message)
                        status_text.update()

            except Exception as driver_ex:
                error_message = f"ドライバ処理中にエラーが発生しました: {str(driver_ex)}\n" + traceback.format_exc()
                status_text.value += error_message
                print(error_message)
                status_text.update()
            finally:
                driver.quit()

            if i < update_count - 1:  # 最後の更新の後には待機しない
                status_text.value += f"{update_interval} 分間待機\n"
                status_text.update()
                print(f"{update_interval} 分間待機")
                progress_bar.visible = True  # プログレスバーを表示
                for j in range(update_interval * 60):
                    progress_bar.value = j / (update_interval * 60)
                    progress_bar.update()
                    time.sleep(1)
                progress_bar.visible = False  # プログレスバーを非表示

        status_text.value += "ツール終了\n"
        status_text.text_style = ft.TextStyle(size=20, weight="bold")
        status_text.update()
        print("ツール終了")

    except ValueError as value_ex:
        error_message = f"無効な商品数、更新間隔、更新回数: {str(value_ex)}\n" + traceback.format_exc()
        status_text.value += error_message
        print(error_message)
    except Exception as ex:
        error_message = f"エラー: {str(ex)}\n" + traceback.format_exc()
        status_text.value += error_message
        print(error_message)
    status_text.update()

def start_refresh(e, status_text, progress_bar, product_count_value, update_interval_value, update_count_value):
    threading.Thread(target=execute_task,
                     args=(status_text, progress_bar, product_count_value, update_interval_value, update_count_value)).start()

def stop_refresh(e, status_text):
    status_text.value += "ツール停止\n"
    status_text.update()
    print("ツール停止")
    ft.app().shutdown()  # ツールを終了するために追加

def main(page: ft.Page):
    page.title = "Minne Refresh Tool"
    page.window_width = 600  # 横幅を倍に設定
    page.window_height = 800  # ウィンドウの高さを調整

    # 全角数値を半角数値に変換する関数
    def convert_to_halfwidth(e):
        e.control.value = e.control.value.translate(str.maketrans(
            '０１２３４５６７８９', '0123456789'))
        e.control.update()

    # 商品数
    product_count = ft.TextField(label="更新商品数", width=400, on_change=convert_to_halfwidth)  # 横幅を倍に設定

    # 更新間隔
    update_interval = ft.TextField(label="更新間隔 (分)", width=400, on_change=convert_to_halfwidth)  # 横幅を倍に設定

    # 更新回数
    update_count = ft.TextField(label="更新回数", width=400, on_change=convert_to_halfwidth)  # 横幅を倍に設定

    # 実行状況表示エリア
    status_text = ft.TextField(value="", width=500, height=300, read_only=True, multiline=True)
    status_text.scroll = "always"

    # プログレスバー
    progress_bar = ft.ProgressBar(width=500, height=20, visible=False)  # 初期状態は非表示

    # スタートボタン
    start_button = ft.ElevatedButton(
        text="スタート",
        on_click=lambda e: start_refresh(e, status_text, progress_bar, product_count.value, update_interval.value, update_count.value),
        style=ft.ButtonStyle(bgcolor=ft.colors.BLUE, color=ft.colors.WHITE)
    )

    # ストップボタン
    stop_button = ft.ElevatedButton(
        text="ストップ",
        on_click=lambda e: stop_refresh(e, status_text),
        style=ft.ButtonStyle(bgcolor=ft.colors.RED, color=ft.colors.WHITE)
    )

    # レイアウト設定
    page.add(
        ft.Column([
            product_count,
            update_interval,
            update_count,
            ft.Row([start_button, stop_button], alignment=ft.MainAxisAlignment.CENTER),
            status_text,
            progress_bar
        ], alignment=ft.MainAxisAlignment.CENTER, spacing=10)
    )

ft.app(target=main)
