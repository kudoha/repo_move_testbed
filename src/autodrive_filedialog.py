from msilib.schema import Error
import uiautomation as uia
import time


class FileDlg:
    """windows環境のファイルダイアログを操作する

            Args:
                
            Raises:
                Exception: 操作中にエラーが発生した場合
    """
    # コンストラクタ
    def __init__(self):
        self.name = "開く"
        self.filedlg = None

    # 開く
    def open(self, path, edit_aotoid="1148", btn_autoid="1"):

        # ファイルダイアログを探す
        w = uia.GetRootControl()
        start_time = time.time()
        while time.time() - start_time < 10:
            time.sleep(0.1)
            try:
                self.filedlg = w.Control(Name=self.name)
            except Error as e:
                raise Exception(f"{str(e)}")
       
        time.sleep(0.1)

        # ファイル名を設定する
        self.filedlg.EditControl(
            AutomationId=edit_aotoid).GetValuePattern().SetValue(path)
        time.sleep(0.1)

        # ボタンをクリックする
        self.filedlg.ButtonControl(
            AutomationId=btn_autoid).Click()
