from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import autodrive_filedialog
from selenium.webdriver.common.alert import Alert
from msilib.schema import Error
import time
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager

class TransPrint:

    def __init__(self, logger=None):
        self._driver = None
        self.logger = logger
    def output_log(self,text):
        if self.logger:
            self.logger.debug(text)
        
    def accept_alert(self):
        """アラートを閉じる

        Args:

        Raises:
            Exception: エラー発生
        """
        try:
            self.output_log("アラート出現待ち")
            WebDriverWait(self._driver, 600).until(EC.alert_is_present())
            Alert(self._driver).accept()
            self.output_log("アラートを消した")
        except TimeoutException:
            pass    # タイムアウトは無視する
        except Exception as e:
            raise Exception(f"Click operation failed: {str(e)}")
        
    def wait_for_element_to_be_visible(self, by, value, timeout=10):
        """指定した要素が表示するまで待つ

        Args:
            by: 要素を検索する方法 (e.g., By.ID, By.XPATH)
            value: 要素のセレクタ値
            timeout: 待つ最大時間 (秒)

        Raises:
            TimeoutException: タイムアウト時に発生
        """
        try:
            self.output_log("{}出現待ち".format(value))
            WebDriverWait(self._driver, timeout).until(
                EC.visibility_of_element_located((by, value))
            )
            self.output_log("{}消失".format(value))
        except TimeoutException:
            raise TimeoutException(f"要素が {timeout} 秒以内に表示されませんでした。")
        
    def wait_for_element_to_be_clickable(self, by, value, timeout=10):
        """指定した要素がクリック可能になるまで待つ

        Args:
            by: 要素を検索する方法 (e.g., By.ID, By.XPATH)
            value: 要素のセレクタ値
            timeout: 待つ最大時間 (秒)

        Raises:
            TimeoutException: タイムアウト時に発生
        """
        try:
            self.output_log("{}活性待ち".format(value))
            WebDriverWait(self._driver, timeout).until(
                EC.element_to_be_clickable((by, value))
            )
            self.output_log("{}活性化".format(value))
        except TimeoutException:
            raise TimeoutException(f"要素が {timeout} 秒以内にクリック可能になりませんでした。")
        
    def click_element(self, by, value):
        """指定したセレクタ（by）および値（value）に一致する要素をクリックします。

        Args:
            by (selenium.webdriver.common.by.By): 要素を検索するための方法 (e.g., By.ID, By.XPATH)
            value (str): 要素のセレクタ値

        Raises:
            Exception: クリック操作が失敗した場合
        """
        try:
            self.output_log("{}クリック開始".format(value))
            element = WebDriverWait(self._driver, 10).until(EC.element_to_be_clickable((by, value)))
            element.click()
            self.output_log("{}クリック終了".format(value))
        except Exception as e:
            raise Exception(f"Click operation failed: {str(e)}")

    def wait_for_loading_bar_to_disappear(self, timeout=60):
        """ローディングバーが非表示になるまで待つ

        Args:
            
        Raises:
            Exception: エラーが発生した場合
        """
        try:
            self.output_log("ローディングバー消失待ち")
            WebDriverWait(self._driver, timeout).until(EC.invisibility_of_element_located((By.CLASS_NAME, "nuxt-progress")))
            self.output_log("ローディングバー消失")
        except Exception as e:
            raise Exception(f"Nuxt Progressが非表示にならないか、タイムアウトしました。: {str(e)}")
        
    def wait_for_processing_element_to_disappear(self, by, value, timeout=360):
        """指定した要素が消えるまで待つ

        Args:
            by (selenium.webdriver.common.by.By): 要素を検索するための方法 (e.g., By.ID, By.XPATH)
            value (str): 要素のセレクタ値
            
        Raises:
            Exception: エラーが発生した場合
        """

        try:
            WebDriverWait(self._driver, timeout).until(EC.invisibility_of_element_located((by, value)))
        except Exception as e:
            raise Exception(f"Nuxt Progressが非表示にならないか、タイムアウトしました。: {str(e)}")
        
    def wait_for_processing_icon_to_disappear(self, timeout=360):
        """処理中アイコンが非表示になるまで待つ

        Args:
            
        Raises:
            Exception: エラーが発生した場合
        """

        try:
            WebDriverWait(self._driver, timeout).until(EC.invisibility_of_element_located((By.CLASS_NAME, "loader")))
        except Exception as e:
            raise Exception(f"Nuxt Progressが非表示にならないか、タイムアウトしました。: {str(e)}")
    
    def get_processing_status(self,label):
        """
        指定したラベルを持つ処理ステータス要素を取得します。
        
        Args:
            label (str): ステータスを識別するラベル
        
        Returns:
            list: ステータス要素のリスト
        Raises:
            Exception: エラーが発生した場合
        """        
        status_els = []
        try:
            els = self._driver.find_element(By.XPATH,"/html/body/div[1]/div/div/div/div[2]/div").find_element(By.CLASS_NAME,"notification").find_elements(By.TAG_NAME,"span")
            for el in els:
                if label in el.text:
                    status_els.append(el)
        except Exception as e:
            raise Exception(f"Processing status retrieval failed: {str(e)}")
        return status_els
    
    def start_browser(self,url):
        """ブラウザを起動する

        Args:
            url (str): 表示するURL
        Raises:
            Exception: エラーが発生した場合
        """
        # ブラウザ起動
        try:
            self._driver = webdriver.Chrome(ChromeDriverManager().install())
            self._driver.get(url)
        except AttributeError as e:
            try:
                self._driver = webdriver.Chrome()
                self._driver.get(url)
            except Exception as e:
                raise Exception(f"{str(e)}")
        except Exception as e:
            raise Exception(f"{str(e)}")

    def login(self,id,password):
        """ログイン操作を行う

        Args:
            id (str): ユーザーID
            password (str): パスワード
        Raises:
            Exception: ログイン時にエラーが発生した場合
        """
        try:
            # id
            self._driver.find_element(By.ID,"userid").send_keys(id)

            # pass
            self._driver.find_element(By.ID,"password").send_keys(password)

            # submit
            self.click_element(By.ID, "submit")

            # 画面が表示するまで待つ
            self.wait_for_element_to_be_visible(By.ID, "logo",60)

        except Exception as e:
            raise Exception(f"{str(e)}")
    
    def open_cloud_settings(self):
        """クラウド設定画面を開く

        Args:
            
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        try:
            # 画面が操作可能になるまで待つ
            self.wait_for_element_to_be_clickable(By.XPATH,"/html/body/viewer-pre-load/div/spa-main/div/div[3]/div[1]/div[1]/div[4]/div/div/button")
            
            # adminボタンをクリック
            self.click_element(By.XPATH, "/html/body/viewer-pre-load/div/spa-main/div/div[3]/div[1]/div[1]/div[4]/div/div/button")

            # クラウド設定項目をクリック
            self.click_element(By.XPATH, "/html/body/viewer-pre-load/div/spa-main/div/div[3]/div[1]/div[1]/div[4]/div/div/ul/li[3]/a")

            # 新しく開いたtabを操作可能にする
            self._driver.switch_to.window(self._driver.window_handles[-1])  # 新しいタブに切り替え
            self._driver.get(self._driver.current_url)

            # 新しく開いたタブの画面が操作可能になるまで待つ
            self.wait_for_element_to_be_visible(By.XPATH,"/html/body/div[1]/div/div/nav/div[1]/a/img")
        except Exception as e:
            raise Exception(f"{str(e)}")
    
    def select_delivery_setting(self):
        """配信先設定を開く

        Args:
            
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        try:
            # 画面の表示が落ち着くまでまつ
            self.wait_for_element_to_be_clickable(By.XPATH,"/html/body/div[1]/div/div/div/div[1]/div/aside/ul/div/li[2]/ul/li[4]/a",60)

            # メニューから配信先管理をクリックする
            self.click_element(By.XPATH,"/html/body/div[1]/div/div/div/div[1]/div/aside/ul/div/li[2]/ul/li[4]/a")

            # 画面の表示が落ち着くまでまつ
            self.wait_for_loading_bar_to_disappear()
        except Exception as e:
            raise Exception(f"{str(e)}")
    
    def click_upload_button(self):
        """アップロードボタンをクリックする

        Args:
            
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        max_retries = 3
        for retry in range(max_retries):
            try:
                # 画面からアップロードボタンをクリックする
                self.click_element(By.XPATH,"/html/body/div[1]/div/div/div/div[2]/div/div[6]/div[1]/button[6]")

                # 画面の表示が落ち着くまでまつ
                self.wait_for_element_to_be_visible(By.CLASS_NAME,"modal-card")
                break   # ループを抜ける
            except Exception as e:
                error_message = str(e)
                if "Click operation failed: Message:" not in error_message:
                    raise Exception(f"Selection of all customer codes failed: {error_message}")
                
                if retry < max_retries - 1:
                    # リトライの前に一時停止する場合はここで time.sleep() を使用できます
                    continue  # リトライ
                else:
                    raise Exception(f"アップロードボタンのクリックに失敗した:{str(e)}")
    def select_filetype(self,filetype):
        """アップロードするファイルのタイプを選択する

        Args:
            filetype (str): アップロードするファイルのタイプ（csv,excel）          
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        try:
            els = self._driver.find_element(By.CLASS_NAME,"modal-card").find_elements(By.TAG_NAME,"input")
            for el in els:
                label = el.get_attribute('value').lower()
                if filetype.lower() in label:
                    el.click()
                    break   # ループを抜ける
        except Exception as e:
            raise Exception(f"アップロードするファイルのタイプの選択に失敗した{str(e)}")

    def clike_select_file_button(self):
        """ファイル選択ボタンをクリックする

        Args:
                    
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        try:
            els = self._driver.find_element(By.CLASS_NAME,"modal-card").find_elements(By.TAG_NAME,"label")
            for el in els:
                e = el.find_elements(By.CLASS_NAME,"file-label")
                if len(e) > 0:
                    e[0].click()
                    break  # ループを抜ける
        except Exception as e:
            raise Exception(f"ファイル選択ボタンのクリックに失敗した{str(e)}")

    def click_cancel_in_form(self):
        """キャンセルボタンをクリックする

        Args:
                    
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        try:
            els = self._driver.find_element(By.CLASS_NAME,"modal-card").find_elements(By.TAG_NAME,"button")
            for el in els:
                if "cancel" in el.text:
                    el.click()
                    break  # ループを抜ける
        except Exception as e:
            raise Exception(f"キャンセルボタンのクリックに失敗した{str(e)}")
    def click_upload_button_in_form(self):
        """アップロードボタンをクリックする

        Args:
                    
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        try:
            els = self._driver.find_element(By.CLASS_NAME,"modal-card").find_elements(By.TAG_NAME,"button")
            for el in els:
                if "アップロード" in el.text:
                    el.click()
                    break  # ループを抜ける
        except Exception as e:
            raise Exception(f"アップロードボタンのクリックに失敗した{str(e)}")

    def clike_check_button(self):
        """チェックボタンをクリックする

        Args:
               
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        try:
            els = self._driver.find_element(By.CLASS_NAME,"modal-card").find_elements(By.TAG_NAME,"button")
            for el in els:
                if "チェック" in el.text:
                    el.click()
                    break  # ループを抜ける
        except Exception as e:
            raise Exception(f"チェックボタンのクリックに失敗した{str(e)}")

    def perform_file_upload(self,filepath,filetype):
        """配信先画面からアップロード操作を行う

        Args:
            filepath (str): アップロードするファイルのパス
            filetype (str): アップロードするファイルのタイプ（csv,excel）   
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        try:
            # 画面からアップロードボタンをクリックする
            self.click_upload_button()
            self.output_log("画面からアップロードボタンをクリックした")

            # アップロードするファイルのタイプを選択する
            self.select_filetype(filetype)
            self.output_log("アップロードするファイルのタイプを選択した")

            # ファイル選択ボタンをクリックする
            self.clike_select_file_button()
            self.output_log("ファイル選択ボタンをクリックした")
            
            # ファイルダイアログからファイルを開く
            autodrive_filedialog.FileDlg().open(filepath)
            self.output_log("ファイルダイアログからファイルを開いた")
           
            # チェックボタンをクリックする
            self.clike_check_button()
            self.output_log("チェックボタンをクリックした")

            # テキストエリアに結果が出るまで待つ
            els = self._driver.find_elements(By.CLASS_NAME, "result-text")
            WebDriverWait(self._driver, 600).until(
                EC.text_to_be_present_in_element_value((By.CLASS_NAME, "result-text"), "成功")
            )
            self.output_log("テキストエリアに結果が出た")

            
            # アップロードをクリックする
            self.click_upload_button_in_form()
            self.output_log("アップロードをクリックした")

            # アラートを閉じる
            self.accept_alert()

        except Exception as e:
            raise Exception(f"{str(e)}")
    
    def select_all_customer_codes(self):
        """ページ内のお客様コードを全件選択状態にする操作を行う

        Args:
            
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        max_retries = 3
        for retry in range(max_retries):
            try:
                
                self.wait_for_processing_icon_to_disappear(600)

                # テーブルのヘッダ行のチェックボックスを探してチェックを入れる
                table_element = self._driver.find_element(By.XPATH,"/html/body/div[1]/div/div/div/div[2]/div").find_element(By.CLASS_NAME,"table")
                checkbox_element = table_element.find_element(By.TAG_NAME,"thead").find_element(By.TAG_NAME,"tr").find_element(By.TAG_NAME,"th").find_element(By.TAG_NAME,"input")
                checkbox_element.click()
                # チェックボックスがONになっていなければリトライ
                time.sleep(0.5)
                is_checkbox = checkbox_element.is_selected()
                if is_checkbox == False:
                    continue
                # 成功時の処理を追加することを検討
                break  # 成功したらループを抜ける
            except Exception as e:
                error_message = str(e)
                if "Click operation failed: Message:" not in error_message:
                    raise Exception(f"Selection of all customer codes failed: {error_message}")
                
                if retry < max_retries - 1:
                    # リトライの前に一時停止する場合はここで time.sleep() を使用できます
                    continue  # リトライ
                else:
                    raise Exception(f"Selection of all customer codes failed after {max_retries} retries.")

        
    def click_modalcard_element(self, by, selector, label):
        """
        指定した要素を探してクリックします。

        Args:
            by: セレクタの種類（例: By.ID, By.NAME）
            selector: セレクタの値
            label: クリックしたい要素に含まれるテキスト

        Raises:
            Exception: エラーが発生した場合
        """

        # modal-cardの要素が表示するまで待つ
        self.wait_for_element_to_be_visible(By.CLASS_NAME, "modal-card")
 
        max_retries = 3
        retries = 0

        while retries < max_retries:
            try:
                els = self._driver.find_element(By.CLASS_NAME, "modal-card").find_elements(by, selector)
                for el in els:
                    if label in el.text.strip():
                        el.click()
                        break  # ループを抜ける
            except ElementClickInterceptedException as e:
                # element click intercepted エラーが発生した場合
                if retries < max_retries - 1:
                    retries += 1
                    # エラーが発生した場合、一時的に待機してから再試行
                    time.sleep(1)
                else:
                    # 最大リトライ回数に達した場合、エラーを再スロー
                    raise Exception("Maximum retries reached, unable to click the element.")
            except Exception as e:
                # その他のエラーが発生した場合は例外をスロー
                raise Exception(f"{str(e)}")
            else:
                # クリックが成功した場合、ループを終了
                break  
    def wait_for_delete_button_to_become_clickable(self):
        """選削除ボタンがクリック可能になるまで画面を更新して待つ

        Args:
            
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        max_retries = 10  # リトライの最大回数
        for retry in range(max_retries + 1):
            try:
                # 更新ボタンを押下する
                self._driver.find_element(By.CLASS_NAME,"table-outer").find_element(By.TAG_NAME,"button").click()
                
                # グルグルが消えるまで待つ
                self.wait_for_processing_icon_to_disappear(60)
                
                # 処理状況を取得する
                els = self.get_processing_status("配信先の削除状況")
                
                # 表示中の処理状況がすべて「処理完了」になっているか確認する
                driver = self._driver
                WebDriverWait(driver, 60).until(lambda driver: all("処理完了" in element.text for element in els))
                print("削除処理完了 - OK")

                return # 処理を終了する

            except TimeoutException:
                if retry < max_retries:
                    # リトライ可能なエラーの場合、次のリトライに進む
                    continue
                else:
                    # 最終リトライ時にもボタンが非活性なら例外をスロー
                    raise Exception(f"{str(e)}")    
            except Exception as e:
                raise Exception(f"{str(e)}")     
                
    def wait_for_upload_button_to_become_clickable(self):
        """アップロードボタンがクリック可能になるまで画面を更新して待つ

        Args:
            
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        max_retries = 100  # リトライの最大回数
        for retry in range(max_retries + 1):
            try:
                # ブラウザの更新
                self.output_log("ブラウザの更新")
                self._driver.refresh()
                self.wait_for_loading_bar_to_disappear(360)
                
                # 画面が落ち着くまで待つ
                self.wait_for_element_to_be_visible(By.XPATH,"/html/body/div[1]/div/div/div/div[2]/div/div[10]/div/div[1]/div/button/img",360)
                self.wait_for_element_to_be_visible(By.CLASS_NAME,"notification",600)
                
                # 処理状況を取得する
                els = self.get_processing_status("配信先の更新状況")
                self.output_log("処理状況を取得した")
                
                # 表示中の処理状況がすべて「処理完了」になっているか確認する
                driver = self._driver
                WebDriverWait(driver, 360).until(lambda driver: all("処理完了" in element.text for element in els))
                self.output_log("配信先の更新完了 - OK")

                return # 処理を終了する

            except TimeoutException:
                if retry < max_retries:
                    # リトライ可能なエラーの場合、次のリトライに進む
                    continue
                else:
                    # 最終リトライ時にもボタンが非活性なら例外をスロー
                    raise Exception(f"{str(e)}")    
            except Exception as e:
                raise Exception(f"{str(e)}")     
    def delete_customer_codes(self):
        """選択中のお客様コードを削除する操作を行う

        Args:
            
        Raises:
            Exception: 操作中にエラーが発生した場合
        """
        try:

            # 削除ボタンをクリックする            
            self.click_element(By.XPATH,"/html/body/div[1]/div/div/div/div[2]/div/div[5]/div[1]/button[5]")

            # 選択した配信先を削除する をクリックする
            self.click_modalcard_element(By.TAG_NAME,"input","")
            print("選択した配信先を削除する をクリックする - OK")

            # 削除ボタンをクリックする
            self.click_modalcard_element(By.TAG_NAME,"button","削除")
            print("削除 をクリックする - OK")     
            
            # ぐるぐるアイコンが消えるまで待つ
            self.wait_for_processing_icon_to_disappear()
            print("ぐるぐるアイコンが消えるまで待つ - OK")

            # 画面の表示が落ち着くまで待つ
            self.wait_for_element_to_be_visible(By.XPATH,"/html/body/div[1]/div/div/div/div[2]/div/div[5]/div[1]/button[5]", 600) 
            print("画面の表示が落ち着くまで待つ - OK")    

            # ぐるぐるアイコン(ローディング)が消えるまで待つ
            self.wait_for_loading_bar_to_disappear()
            print("ぐるぐるアイコン(ローディング)が消えるまで待つ - OK")   
       
        except Exception as e:
            raise Exception(f"{str(e)}")          


        

        
        