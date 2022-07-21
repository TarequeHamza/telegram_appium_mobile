import time

from Pages.BasePage import BasePage
from Pages.TelegramPage import TelegramPage


class LoginTest(BasePage):

    def test_LoginPage(self):
        login = TelegramPage(self.driver)
        login.click_startmessaging()
        time.sleep(3)

