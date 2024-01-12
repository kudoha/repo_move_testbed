import unittest
import time
from unittest.mock import MagicMock
from autodrive_transprint import TransPrint  # テスト対象のモジュールをインポート

class TestTransPrint(unittest.TestCase):

    def test(self):
        base_url = "https://qa224.dev.spa-cloud.com/spa/"
        user_id = "spa"
        password = "spa"
        transprint = TransPrint()
        transprint.start_browser(base_url)
        transprint.login(user_id,password)
        transprint.open_cloud_settings()
        transprint.select_delivery_setting()
        time.sleep(5)
        


if __name__ == '__main__':
    unittest.main()