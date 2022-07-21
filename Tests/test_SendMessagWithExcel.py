from Pages.BasePage import BasePage
from Pages.SendMessagePage import SendMessagePage
from Pages.ContactSavePage import SaveContactPage
from utils.ExcelHandling import *
import os.path


class SendMessageTest(BasePage):
    def test_SendMessage(self):
        send = SendMessagePage(self.driver)
        save = SaveContactPage(self.driver)
        excel = os.path.abspath('..\TestData\Contract_Number_List.xlsx')
        firstName = open_and_read_excel_file(excel, "Sheet1", 1)
        lastName = open_and_read_excel_file(excel, "Sheet1", 2)
        number = open_and_read_excel_file(excel, "Sheet1", 3)
        for i in range(len(number)):
            send.click_message_icon()
            send.click_search_icon()
            send.type_phone_number(number[i])
            Text = f"No results found for '{number[i]}'"
            if Text == save.get_text():
                save.click_backArrow_icon()
                save.click_new_contact_button()
                save.type_first_name(firstName[i])
                save.type_last_name(lastName[i])
                save.type_phone_number(number[i])
                save.click_save_button()
                save.click_backArrow_icon()

            else:
                send.click_expected_profile()
                send.type_message(f"Hi {firstName[i]} {lastName[i]},"
                                    "\n\nNext Friday And Saturday Our Office Will Remain Closed, Enjoy Your Time!,"
                                    "\n\nHR, QUPS")
                send.click_send_button()
                writeData(excel, "Sheet1", i + 2, 4, "1")
                send.click_back_button()
