import time

from Pages.BasePage import BasePage
from Pages.SendMessagePage import SendMessagePage
 

class SendMessageTest(BasePage):

    def test_SendMessage(self):
        send = SendMessagePage(self.driver)
        send.click_message_icon()
        send.click_search_icon()
        send.type_phone_number("01611844007")
        send.click_expected_profile()
        send.type_message("Hi, Meraz, Welcome To You From QUPS")
        send.click_send_button()

